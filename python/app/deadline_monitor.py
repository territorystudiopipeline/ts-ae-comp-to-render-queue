# Copyright (c) 2026 Territory Studio
"""
Links render queue items to their Deadline jobs and follows those jobs.

The link lives in the render queue item's Comment (shown in After Effects' Render Queue panel,
saved with the project), as a tag such as "[Deadline 6a3f19c0d1e2f3a4b5c6d7e8 queued]". Other
comment text is kept. After Effects can't mark an item Done from a script, so a submitted item
is unqueued and the tag records the job and its last known state; the table builds farm rows
from it whenever the window opens.

DeadlineJobMonitor asks the Deadline Web Service (DeadlineConnect) about all followed jobs in
one request, on a worker thread, every POLL_MS while the window is open.
"""
import re

import sgtk
from sgtk.platform.qt import QtCore

logger = sgtk.platform.get_logger(__name__)

# Job states (also written in the comment tag)
QUEUED = "queued"
RENDERING = "rendering"
SUSPENDED = "suspended"
COMPLETED = "completed"
FAILED = "failed"
NOT_FOUND = "missing"
UNKNOWN = "unknown"
STATES = (QUEUED, RENDERING, SUSPENDED, COMPLETED, FAILED, NOT_FOUND, UNKNOWN)

TAG_RE = re.compile(r"\s*\[Deadline ([0-9A-Za-z]+)(?: ([a-z]+))?\]")


# ---------------------------------------------------------------------------------- comment tag

def read_tag(comment):
    """
        The Deadline link in a render queue item comment.

        Returns:
            tuple: (job_id, state), or None when the comment has no tag.
    """
    match = TAG_RE.search(comment or "")
    if not match:
        return None
    state = match.group(2) if match.group(2) in STATES else UNKNOWN
    return match.group(1), state


def write_tag(comment, job_id, state):
    """
        The comment with its Deadline tag set (added at the end, or replaced in place).
    """
    tag = "[Deadline %s %s]" % (job_id, state)
    comment = comment or ""
    if TAG_RE.search(comment):
        return TAG_RE.sub(lambda m: (" " if m.start() else "") + tag, comment, count=1)
    return ("%s %s" % (comment.rstrip(), tag)).strip()


def clear_tag(comment):
    """The comment without its Deadline tag."""
    return TAG_RE.sub("", comment or "").strip()


# ---------------------------------------------------------------------------------- job status

# Deadline job "Stat" values
_STAT = {0: UNKNOWN, 1: QUEUED, 2: SUSPENDED, 3: COMPLETED, 4: FAILED, 6: QUEUED}


def job_info(job):
    """
        Summarise a Deadline job (a DeadlineConnect job dict).

        Returns:
            dict: {"state", "done", "total", "failed_tasks", "name"}
    """
    def number(key):
        try:
            return int(job.get(key) or 0)
        except (TypeError, ValueError):
            return 0

    props = job.get("Props") or {}
    state = _STAT.get(number("Stat"), UNKNOWN)
    done = number("CompletedChunks")
    total = 0
    try:
        total = int(props.get("Tasks") or 0)
    except (TypeError, ValueError):
        pass
    if not total:
        total = sum(number(key) for key in ("CompletedChunks", "QueuedChunks", "RenderingChunks", "FailedChunks",
                                            "PendingChunks", "SuspendedChunks"))
    if state == QUEUED and number("RenderingChunks"):
        state = RENDERING
    elif state == QUEUED and done:
        state = RENDERING  # between tasks: some done, more to go
    return {"state": state, "done": done, "total": total, "failed_tasks": number("FailedChunks"),
            "name": props.get("Name") or ""}


def fetch_jobs(host, port, job_ids):
    """
        Ask the Deadline Web Service about some jobs. Safe in any thread.

        Returns:
            dict: {job_id: job_info(...)}; jobs Deadline doesn't know are NOT_FOUND.

        Raises:
            Exception: If Deadline can't be reached.
    """
    import Deadline.DeadlineConnect as Connect

    deadline = Connect.DeadlineCon(host, port)
    job_ids = list(job_ids)
    jobs = deadline.Jobs.GetJobs(job_ids)
    if not isinstance(jobs, list):
        # DeadlineConnect returns the error text as a string
        raise Exception(str(jobs))
    result = {}
    for job in jobs:
        if isinstance(job, dict) and job.get("_id") in job_ids:
            result[job["_id"]] = job_info(job)
    for job_id in job_ids:
        result.setdefault(job_id, {"state": NOT_FOUND, "done": 0, "total": 0, "failed_tasks": 0, "name": ""})
    return result


class _FetchWorker(QtCore.QObject):
    """Runs fetch_jobs on the monitor's thread; talks to the GUI only through signals."""
    fetched = QtCore.Signal(object, str)  # {job_id: info}, error text

    def __init__(self, host, port):
        super(_FetchWorker, self).__init__()
        self.host, self.port = host, port

    @QtCore.Slot(object)
    def fetch(self, job_ids):
        try:
            self.fetched.emit(fetch_jobs(self.host, self.port, job_ids), "")
        except Exception as e:
            self.fetched.emit({}, str(e) or e.__class__.__name__)


# Threads still finishing a request when their monitor was stopped (a QThread destroyed while
# running aborts the process, and a Web Service request can't be interrupted)
_FINISHING = set()


class DeadlineJobMonitor(QtCore.QObject):
    """
        Follow Deadline jobs while the window is open.

        Usage:
            monitor = DeadlineJobMonitor(host, port, parent=dialog)
            monitor.job_updated.connect(...)
            monitor.set_jobs(["6a3f..."])   # starts polling

        Signals:
            job_updated (str, dict): A job's latest info (see job_info), emitted on every poll.
            error (str): Deadline couldn't be reached (polling carries on).
    """
    job_updated = QtCore.Signal(str, object)
    error = QtCore.Signal(str)

    POLL_MS = 20000
    # Jobs in these states are no longer polled
    FINAL_STATES = (COMPLETED, NOT_FOUND)

    _fetch_requested = QtCore.Signal(object)

    def __init__(self, host, port, parent=None):
        super(DeadlineJobMonitor, self).__init__(parent)
        self.host, self.port = host, port
        self._jobs = set()
        self._in_flight = False
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.poll_now)
        self._thread = None
        self._worker = None

    def set_jobs(self, job_ids):
        """Follow these jobs (replaces the list). Polls straight away when it changes."""
        job_ids = set(job_ids)
        added = job_ids - self._jobs
        self._jobs = job_ids
        if not self._jobs:
            self._timer.stop()
        elif added:
            self.poll_now()

    def add_job(self, job_id):
        self.set_jobs(self._jobs | {job_id})

    def jobs(self):
        return set(self._jobs)

    def poll_now(self):
        if not self._jobs:
            return
        if self._in_flight:
            return  # the reply schedules the next poll
        self._start_thread()
        self._in_flight = True
        self._fetch_requested.emit(sorted(self._jobs))

    def stop(self):
        """Stop polling. A request already sent finishes on its own."""
        self._jobs = set()
        self._timer.stop()
        thread, self._thread = self._thread, None
        if thread is None:
            return
        try:
            self._fetch_requested.disconnect()
        except (RuntimeError, TypeError):
            pass
        thread.quit()
        if not thread.wait(2000):
            # Still waiting on Deadline: keep the thread alive until it ends by itself
            thread.setParent(None)
            _FINISHING.add(thread)
            thread.finished.connect(lambda t=thread: _FINISHING.discard(t))
            logger.debug("Deadline status request still running; leaving it to finish")

    # ------------------------------------------------------------------ internals

    def _start_thread(self):
        if self._thread is not None:
            return
        self._thread = QtCore.QThread()
        self._worker = _FetchWorker(self.host, self.port)
        self._worker.moveToThread(self._thread)
        self._fetch_requested.connect(self._worker.fetch, QtCore.Qt.QueuedConnection)
        self._worker.fetched.connect(self._on_fetched, QtCore.Qt.QueuedConnection)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.start()

    def _on_fetched(self, results, error):
        self._in_flight = False
        if self._thread is None:
            return  # stopped meanwhile
        if error:
            logger.debug("Deadline status check failed: %s" % error)
            self.error.emit(error)
        for job_id, info in results.items():
            if job_id not in self._jobs:
                continue
            if info["state"] in self.FINAL_STATES:
                self._jobs.discard(job_id)
            self.job_updated.emit(job_id, info)
        if self._jobs:
            self._timer.start(self.POLL_MS)
