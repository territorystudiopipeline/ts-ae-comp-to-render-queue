# Copyright (c) 2026 Territory Studio
"""
Follows a local After Effects render (renderQueue.renderAsync()) so the table can update live.

LocalRenderMonitor polls the render queue through the tk-aftereffects bridge on a GUI-thread
timer (the bridge is not thread-safe). Measured in After Effects 26.5: renderAsync() returns
after ~0.6s, and a status query takes ~4ms during a render, so a 1s poll is cheap.

RenderQueueItem.elapsedSeconds stays 0 while an item renders, so progress for image sequences
comes from FrameCountWorker instead: it counts the frames written since the render started
(new files, or files whose modification time changed, compared with a listing taken before
the render) on a worker thread, so a slow network share never holds up the window.
Movies have no frames to count and only show Rendering, then Done.
"""
import os
import re
import time

import sgtk
from sgtk.platform.qt import QtCore

logger = sgtk.platform.get_logger(__name__)

SEQUENCE_TOKEN = re.compile(r"\[(#+)\]")

# Monitor states for one render queue item
WAITING = "waiting"      # queued, the render hasn't reached it yet
RENDERING = "rendering"
DONE = "done"
STOPPED = "stopped"      # stopped by the user
FAILED = "failed"        # stopped by an error
NOT_RENDERED = "not_rendered"  # still queued when the render ended
FINAL_STATES = (DONE, STOPPED, FAILED, NOT_RENDERED)


def sequence_pattern(path):
    """
        (folder, compiled regex) for an image sequence output path such as
        "C:/renders/shot_[####].exr", or None for a single-file (movie) output.
    """
    if not path:
        return None
    folder, name = os.path.split(path)
    match = SEQUENCE_TOKEN.search(name)
    if not match:
        return None
    pattern = re.compile("^%s\\d{%d,}%s$" % (
        re.escape(name[:match.start()]), len(match.group(1)), re.escape(name[match.end():])), re.IGNORECASE)
    return folder, pattern


def list_sequence(path):
    """
        {file name: modification time} for the frames of a sequence output path on disk.
        Empty when the folder doesn't exist yet; None for a movie path.
    """
    parsed = sequence_pattern(path)
    if parsed is None:
        return None
    folder, pattern = parsed
    frames = {}
    try:
        for entry in os.scandir(folder):
            if pattern.match(entry.name):
                try:
                    frames[entry.name] = entry.stat().st_mtime
                except OSError:
                    pass
    except OSError:
        pass
    return frames


def count_written_frames(path, baseline):
    """
        Frames of a sequence written since `baseline` (a list_sequence() result) was taken:
        new files, and existing files whose modification time changed (re-rendering over an
        existing version). Compares file server times with file server times, so clock
        differences between this machine and the server don't matter.
    """
    current = list_sequence(path)
    if current is None:
        return None
    baseline = baseline or {}
    return sum(1 for name, mtime in current.items() if baseline.get(name) != mtime)


class FrameCountWorker(QtCore.QObject):
    """
        Lists output folders on a worker thread. Only talks to the GUI through queued signals.
    """
    counted = QtCore.Signal(object, int)  # key, frames written

    def __init__(self):
        super(FrameCountWorker, self).__init__()
        self._baselines = {}

    @QtCore.Slot(object, object)
    def take_baseline(self, key, paths):
        for path in paths:
            listing = list_sequence(path)
            if listing is not None:
                self._baselines[(key, path)] = listing

    @QtCore.Slot(object, object)
    def count(self, key, paths):
        total = None
        for path in paths:
            frames = count_written_frames(path, self._baselines.get((key, path)))
            if frames is not None:
                total = max(total or 0, frames)  # several sequence outputs write the same frames
        if total is not None:
            self.counted.emit(key, total)


class LocalRenderMonitor(QtCore.QObject):
    """
        Follow a local render and report each item's progress.

        Usage:
            monitor = LocalRenderMonitor(adobe, entries, parent=dialog)
            monitor.item_changed.connect(...)
            monitor.start()

        Arguments:
            adobe: engine.adobe.
            entries (list): One dict per render queue item to follow:
                {"key": row key, "item": RenderQueueItem proxy,
                 "outputs": [output file paths], "expected": frames to render or None}
            count_frames (bool): Count frames on disk. Needs take_baselines() before the render
                starts, so it's off when following a render that was already running.
    """
    item_changed = QtCore.Signal(object, str, str)   # key, state, message
    progress = QtCore.Signal(object, int, int)       # key, frames written, expected (0 = unknown)
    finished = QtCore.Signal(object)                 # summary dict

    POLL_MS = 1000
    # Polls with nothing rendering before the render counts as finished (After Effects can
    # report rendering=False for a moment between items)
    IDLE_POLLS_TO_FINISH = 2

    _baseline_requested = QtCore.Signal(object, object)
    _count_requested = QtCore.Signal(object, object)

    def __init__(self, adobe, entries, count_frames=True, parent=None):
        super(LocalRenderMonitor, self).__init__(parent)
        self.adobe = adobe
        self.entries = []
        for entry in entries:
            entry = dict(entry)
            entry.setdefault("outputs", [])
            entry["sequence"] = [p for p in entry["outputs"] if sequence_pattern(p)] if count_frames else []
            entry["state"] = None
            entry["started"] = None
            entry["seconds"] = None
            entry["frames"] = None
            entry["count_pending"] = False
            self.entries.append(entry)
        self._by_key = dict((entry["key"], entry) for entry in self.entries)

        status = adobe.RQItemStatus
        self._status_states = {
            status.QUEUED: WAITING,
            status.RENDERING: RENDERING,
            status.DONE: DONE,
            status.USER_STOPPED: STOPPED,
            status.ERR_STOPPED: FAILED,
            status.WILL_CONTINUE: DONE,
        }
        self._render_queue = None
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._poll)
        self._in_poll = False
        self._idle_polls = 0
        self._seen_rendering = False
        self._started_at = None
        self.active = False

        self._thread = None
        self._worker = None

    # ------------------------------------------------------------------ control

    def start(self):
        """Start following the render. Call right after renderAsync()."""
        self.active = True
        self._started_at = time.time()
        self._render_queue = self.adobe.app.project.renderQueue
        self._start_frame_counting()
        for entry in self.entries:
            self._set_state(entry, WAITING)
        self._timer.start(0)

    def take_baselines(self):
        """
            List the existing frames now, before the render writes any (call before renderAsync()).
        """
        self._start_frame_counting()

    def stop(self):
        """Stop following (the render itself carries on in After Effects)."""
        self.active = False
        self._timer.stop()
        self._stop_frame_counting()

    def is_running(self):
        return self.active

    # ------------------------------------------------------------------ polling (GUI thread)

    def _poll(self):
        if not self.active or self._in_poll:
            return
        self._in_poll = True
        try:
            self._poll_once()
        except Exception as e:
            logger.warning("Lost track of the render: %s" % e)
            self._finish(error=str(e))
            return
        finally:
            self._in_poll = False
        if self.active:
            self._timer.start(self.POLL_MS)

    def _poll_once(self):
        rendering = bool(self._render_queue.rendering)
        any_rendering = rendering
        for entry in self.entries:
            if entry["state"] in FINAL_STATES:
                continue
            state = self._status_states.get(entry["item"].status)
            if state is None:  # e.g. unqueued while rendering
                state = NOT_RENDERED
            if state == RENDERING:
                any_rendering = True
                self._request_count(entry)
            if state != entry["state"]:
                self._set_state(entry, state)

        if any_rendering:
            self._seen_rendering = True
            self._idle_polls = 0
            return
        self._idle_polls += 1
        if self._idle_polls >= self.IDLE_POLLS_TO_FINISH:
            self._finish()

    def _set_state(self, entry, state):
        previous = entry["state"]
        entry["state"] = state
        now = time.time()
        if state == RENDERING:
            entry["started"] = now
        elif previous == RENDERING and entry["started"] is not None:
            entry["seconds"] = now - entry["started"]
        if state in (DONE, STOPPED, FAILED):
            self._request_count(entry, force=True)  # final frame count
        self.item_changed.emit(entry["key"], state, self.describe(entry))

    def _finish(self, error=None):
        for entry in self.entries:
            if entry["state"] in (WAITING, None):
                self._set_state(entry, NOT_RENDERED)
            elif entry["state"] == RENDERING:
                # Rendering with the render over: After Effects didn't report the end (e.g. the
                # project was closed). Leave it to a refresh.
                self._set_state(entry, NOT_RENDERED)
        counts = {}
        for entry in self.entries:
            counts[entry["state"]] = counts.get(entry["state"], 0) + 1
        summary = {"counts": counts, "seconds": time.time() - (self._started_at or time.time()),
                   "error": error}
        self.active = False
        self._timer.stop()
        # Let the final frame counts arrive before the thread goes away
        QtCore.QTimer.singleShot(1500, self._stop_frame_counting)
        self.finished.emit(summary)

    def describe(self, entry):
        """Short text for an entry's status cell."""
        state = entry["state"]
        if state == RENDERING:
            return self._frames_text(entry) or "Started"
        if state in (DONE, STOPPED, FAILED):
            parts = []
            if entry["seconds"] is not None:
                parts.append(format_seconds(entry["seconds"]))
            frames = self._frames_text(entry)
            if frames:
                parts.append(frames)
            return "  ·  ".join(parts)
        if state == NOT_RENDERED:
            return "Not rendered"
        return ""

    @staticmethod
    def _frames_text(entry):
        if entry["frames"] is None:
            return ""
        if entry.get("expected"):
            return "%d/%d frames" % (min(entry["frames"], entry["expected"]), entry["expected"])
        return "%d frames" % entry["frames"]

    # ------------------------------------------------------------------ frame counting

    def _start_frame_counting(self):
        if self._thread is not None:
            return
        if not any(entry["sequence"] for entry in self.entries):
            return
        self._thread = QtCore.QThread(self)
        self._worker = FrameCountWorker()
        self._worker.moveToThread(self._thread)
        self._baseline_requested.connect(self._worker.take_baseline, QtCore.Qt.QueuedConnection)
        self._count_requested.connect(self._worker.count, QtCore.Qt.QueuedConnection)
        self._worker.counted.connect(self._on_counted, QtCore.Qt.QueuedConnection)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.start()
        for entry in self.entries:
            if entry["sequence"]:
                self._baseline_requested.emit(entry["key"], list(entry["sequence"]))

    def _stop_frame_counting(self, wait_ms=3000):
        thread = self._thread
        if thread is None:
            return
        self._thread = None
        try:
            thread.quit()
            if not thread.wait(wait_ms):
                logger.warning("Frame counting is still listing a folder; leaving it to finish")
        except RuntimeError:
            pass

    def _request_count(self, entry, force=False):
        if self._thread is None or not entry["sequence"]:
            return
        if entry["count_pending"] and not force:
            return
        entry["count_pending"] = True
        self._count_requested.emit(entry["key"], list(entry["sequence"]))

    def _on_counted(self, key, frames):
        entry = self._by_key.get(key)
        if entry is None:
            return
        entry["count_pending"] = False
        entry["frames"] = frames
        self.progress.emit(key, frames, entry.get("expected") or 0)
        if entry["state"] in (RENDERING, DONE, STOPPED, FAILED):
            self.item_changed.emit(key, entry["state"], self.describe(entry))


def format_seconds(seconds):
    seconds = int(round(seconds))
    if seconds < 60:
        return "%ds" % seconds
    minutes, seconds = divmod(seconds, 60)
    if minutes < 60:
        return "%dm %02ds" % (minutes, seconds)
    hours, minutes = divmod(minutes, 60)
    return "%dh %02dm" % (hours, minutes)
