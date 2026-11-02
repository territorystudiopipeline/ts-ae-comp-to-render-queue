# Copyright (c) 2026 Territory Studio
"""
SPIKE (temporary): can the app follow a local render live?

Starts (or watches) an After Effects local render and, on a timer, asks After Effects
for render queue state through the tk-aftereffects bridge (engine.adobe), timing every
call. It can also count rendered frames on disk without asking After Effects anything.

It answers:
    1. Does renderQueue.renderAsync() return straight away, or block until the render ends?
    2. Do status queries answer during a render, and how long do they take?
    3. Does the Toolkit UI stay responsive (timer drift) while After Effects renders?
    4. Is counting frames on disk a usable progress source?

Everything goes to its own log file in the Toolkit log folder (not the Toolkit log, so the
polling doesn't flood After Effects' console). The log path is shown in the window.

Run it:
    - Set TS_AE_RENDER_SPIKE=1 before launching After Effects: the app adds a
      "Render Queue Spike..." command to the ShotGrid menu.
    - Or from a Toolkit Python console in After Effects:
          exec(open(r"<app folder>/python/app/spike_render_poll.py").read()); run()
"""
import datetime
import os
import re
import time

import sgtk
from sgtk.platform.qt import QtCore, QtGui

STATUS_NAMES = ("QUEUED", "NEEDS_OUTPUT", "DONE", "USER_STOPPED", "ERR_STOPPED",
                "RENDERING", "UNQUEUED", "WILL_CONTINUE")
SEQUENCE_TOKEN = re.compile(r"\[(#+)\]")


def _log_folder():
    try:
        return sgtk.LogManager().log_folder
    except Exception:
        import tempfile
        return tempfile.gettempdir()


class _Stats(object):
    """Call timings per label."""

    def __init__(self):
        self.samples = {}

    def add(self, label, seconds):
        self.samples.setdefault(label, []).append(seconds)

    def lines(self):
        result = []
        for label in sorted(self.samples):
            values = sorted(self.samples[label])
            count = len(values)
            p95 = values[min(count - 1, int(round(count * 0.95)) - 1)] if count else 0
            result.append("%-28s n=%-5d min=%7.1fms avg=%7.1fms p95=%7.1fms max=%7.1fms" % (
                label, count, values[0] * 1000, sum(values) / count * 1000, p95 * 1000, values[-1] * 1000))
        return result


class RenderSpikeWidget(QtGui.QWidget):

    def __init__(self, parent=None):
        super(RenderSpikeWidget, self).__init__(parent)
        self.engine = sgtk.platform.current_engine()
        self.adobe = self.engine.adobe
        self.status_names = {}
        self.items = []          # snapshot: dicts per render queue item
        self.stats = _Stats()
        self.polling = False
        self._in_poll = False
        self._last_tick = None
        self._idle_polls = 0
        self._started = None
        self._seen_rendering = False
        self._in_render_call = False

        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = os.path.join(_log_folder(), "ts_ae_render_spike_%s.log" % stamp)
        self._log_file = open(self.log_path, "a")

        self._build_ui()
        self.log("Log file: %s" % self.log_path)
        self.log("After Effects %s, engine %s" % (self._safe(lambda: self.adobe.app.version),
                                                  getattr(self.engine, "version", "?")))
        self.snapshot()

    # ------------------------------------------------------------------ ui

    def _build_ui(self):
        layout = QtGui.QVBoxLayout(self)
        row = QtGui.QHBoxLayout()
        self.snapshot_button = QtGui.QPushButton("Snapshot")
        self.render_button = QtGui.QPushButton("Start renderAsync()")
        self.poll_button = QtGui.QPushButton("Start polling")
        self.summary_button = QtGui.QPushButton("Write summary")
        for button in (self.snapshot_button, self.render_button, self.poll_button, self.summary_button):
            row.addWidget(button)
        layout.addLayout(row)

        options = QtGui.QHBoxLayout()
        self.rpc_check = QtGui.QCheckBox("Ask After Effects (RPC)")
        self.rpc_check.setChecked(True)
        self.disk_check = QtGui.QCheckBox("Count frames on disk")
        self.disk_check.setChecked(True)
        self.interval_spin = QtGui.QSpinBox()
        self.interval_spin.setRange(100, 10000)
        self.interval_spin.setSingleStep(250)
        self.interval_spin.setValue(1000)
        self.interval_spin.setSuffix(" ms")
        options.addWidget(self.rpc_check)
        options.addWidget(self.disk_check)
        options.addStretch()
        options.addWidget(QtGui.QLabel("Interval"))
        options.addWidget(self.interval_spin)
        layout.addLayout(options)

        self.output = QtGui.QPlainTextEdit()
        self.output.setReadOnly(True)
        self.output.setLineWrapMode(QtGui.QPlainTextEdit.NoWrap)
        font = QtGui.QFont("Consolas")
        font.setStyleHint(QtGui.QFont.Monospace)
        self.output.setFont(font)
        layout.addWidget(self.output)
        self.resize(900, 520)

        self.snapshot_button.clicked.connect(self.snapshot)
        self.render_button.clicked.connect(self.start_render)
        self.poll_button.clicked.connect(self.toggle_polling)
        self.summary_button.clicked.connect(self.write_summary)

    def log(self, text):
        line = "%s  %s" % (datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3], text)
        self._log_file.write(line + "\n")
        self._log_file.flush()
        self.output.appendPlainText(line)

    # ------------------------------------------------------------------ rpc helpers

    def timed(self, label, func):
        """Call func, record how long it took. Returns (value, error)."""
        start = time.perf_counter()
        try:
            value, error = func(), None
        except Exception as e:
            value, error = None, e
        elapsed = time.perf_counter() - start
        self.stats.add(label, elapsed)
        if error is not None:
            self.log("  ! %s failed after %.1fms: %s" % (label, elapsed * 1000, error))
        elif elapsed > 0.5:
            self.log("  ~ %s slow: %.1fms" % (label, elapsed * 1000))
        return value, error

    @staticmethod
    def _safe(func, default="?"):
        try:
            return func()
        except Exception:
            return default

    def status_name(self, value):
        return self.status_names.get(value, str(value))

    # ------------------------------------------------------------------ snapshot

    def snapshot(self):
        """Read the render queue once: items, statuses, output paths, expected frames."""
        if not self.status_names:
            for name in STATUS_NAMES:
                value = self._safe(lambda: getattr(self.adobe.RQItemStatus, name), None)
                if value is not None:
                    self.status_names[value] = name
            self.log("RQItemStatus: %s" % ", ".join("%s=%s" % (n, v) for v, n in self.status_names.items()))

        render_queue, error = self.timed("snapshot.renderQueue", lambda: self.adobe.app.project.renderQueue)
        if error:
            return
        count, _ = self.timed("snapshot.numItems", lambda: render_queue.numItems)
        self.items = []
        start = time.perf_counter()
        for index in range(1, (count or 0) + 1):
            item = render_queue.item(index)
            info = {"index": index, "proxy": item,
                    "comp": self._safe(lambda: item.comp.name),
                    "status": self._safe(lambda: item.status, None),
                    "outputs": [], "expected": None, "frames": None}
            duration = self._safe(lambda: item.timeSpanDuration, None)
            frame_rate = self._safe(lambda: item.comp.frameRate, None)
            if duration is not None and frame_rate:
                info["expected"] = int(round(duration * frame_rate))
            for om in range(1, self._safe(lambda: item.numOutputModules, 0) + 1):
                path = self._safe(lambda: item.outputModule(om).file.fsName, None)
                if path:
                    info["outputs"].append(path)
            self.items.append(info)
            self.log("#%d %-30s %-13s frames=%s outputs=%s" % (
                index, info["comp"], self.status_name(info["status"]), info["expected"], info["outputs"]))
        self.log("Snapshot of %s items took %.0fms" % (count, (time.perf_counter() - start) * 1000))

    # ------------------------------------------------------------------ render

    def start_render(self):
        render_queue, error = self.timed("render.renderQueue", lambda: self.adobe.app.project.renderQueue)
        if error:
            return
        queued = [i for i in self.items if self.status_name(i["status"]) == "QUEUED"]
        if not queued:
            self.log("Nothing QUEUED in the snapshot: queue an item in After Effects, then Snapshot.")
            return
        self.log("Calling renderAsync() with %d queued items..." % len(queued))
        self._started = time.perf_counter()
        self.start_polling()  # first poll is scheduled now: shows whether the GUI ran during the call
        # While the call is out, polls skip the bridge (no nested RPC) and only count frames
        self._in_render_call = True
        try:
            _, error = self.timed("renderAsync()", lambda: render_queue.renderAsync())
        finally:
            self._in_render_call = False
        self.log("renderAsync() returned after %.1fms%s" % (
            (time.perf_counter() - self._started) * 1000, " (error)" if error else ""))

    # ------------------------------------------------------------------ polling

    def toggle_polling(self):
        if self.polling:
            self.stop_polling("stopped by user")
        else:
            self.start_polling()

    def start_polling(self):
        if self.polling:
            return
        self.polling = True
        self._idle_polls = 0
        self._seen_rendering = False
        self._last_tick = None
        self.poll_button.setText("Stop polling")
        self.log("Polling every %dms (rpc=%s, disk=%s)" % (
            self.interval_spin.value(), self.rpc_check.isChecked(), self.disk_check.isChecked()))
        QtCore.QTimer.singleShot(0, self.poll)

    def stop_polling(self, reason):
        if not self.polling:
            return
        self.polling = False
        self.poll_button.setText("Start polling")
        self.log("Polling stopped: %s" % reason)
        if self.disk_check.isChecked():
            parts = []
            self._poll_disk(parts)
            self.log("final disk count | %s" % "  ".join(parts))
        self.write_summary()

    def poll(self):
        if not self.polling:
            return
        if self._in_poll:
            # The bridge processes Qt events while waiting: a timer fired inside an RPC call
            self.log("  ! re-entrant poll skipped")
            return
        self._in_poll = True
        now = time.perf_counter()
        drift = None
        if self._last_tick is not None:
            drift = (now - self._last_tick) - self.interval_spin.value() / 1000.0
            self.stats.add("timer drift", max(drift, 0))
        self._last_tick = now
        try:
            parts = []
            rendering = None
            if self._in_render_call:
                parts.append("(GUI ran while renderAsync() is still waiting: RPC skipped)")
            elif self.rpc_check.isChecked():
                rendering = self._poll_rpc(parts)
            if self.disk_check.isChecked():
                self._poll_disk(parts)
            elapsed = (time.perf_counter() - now) * 1000
            self.stats.add("whole poll", elapsed / 1000)
            self.log("poll %5.0fms drift=%s rendering=%s | %s" % (
                elapsed, "-" if drift is None else "%+.0fms" % (drift * 1000), rendering, "  ".join(parts)))
            self._check_finished(rendering)
        finally:
            self._in_poll = False
        if self.polling:
            QtCore.QTimer.singleShot(self.interval_spin.value(), self.poll)

    def _poll_rpc(self, parts):
        render_queue, error = self.timed("rpc renderQueue", lambda: self.adobe.app.project.renderQueue)
        if error:
            return None
        rendering, _ = self.timed("rpc rendering", lambda: render_queue.rendering)
        for info in self.items:
            item = info["proxy"]
            status, error = self.timed("rpc item.status", lambda: item.status)
            if error:
                parts.append("#%d ?" % info["index"])
                continue
            if status != info["status"]:
                self.log("  > #%d %s: %s -> %s" % (info["index"], info["comp"],
                                                  self.status_name(info["status"]), self.status_name(status)))
                info["status"] = status
            text = "#%d %s" % (info["index"], self.status_name(status))
            if self.status_name(status) == "RENDERING":
                elapsed, _ = self.timed("rpc item.elapsedSeconds", lambda: item.elapsedSeconds)
                text += " %ss" % elapsed
            parts.append(text)
        return rendering

    def _poll_disk(self, parts):
        start = time.perf_counter()
        for info in self.items:
            for path in info["outputs"]:
                count = count_sequence_frames(path)
                if count is None:
                    continue
                if count != info["frames"]:
                    info["frames"] = count
                parts.append("#%d %s/%s frames" % (info["index"], count, info["expected"] or "?"))
        self.stats.add("disk count", time.perf_counter() - start)

    def _check_finished(self, rendering):
        """Stop a few polls after the render ends (RPC mode only)."""
        if rendering is None:
            return
        busy = rendering or any(self.status_name(i["status"]) == "RENDERING" for i in self.items)
        if busy:
            self._seen_rendering = True
            self._idle_polls = 0
            return
        if self._seen_rendering or self._started is not None:
            self._idle_polls += 1
            if self._idle_polls >= 3:
                if self._started is not None:
                    self.log("Render finished ~%.1fs after renderAsync() was called"
                             % (time.perf_counter() - self._started))
                    self._started = None
                self.stop_polling("render finished")

    # ------------------------------------------------------------------ summary

    def write_summary(self):
        self.log("---- summary ----")
        for line in self.stats.lines():
            self.log(line)
        for info in self.items:
            self.log("#%d %-30s final=%s frames_on_disk=%s expected=%s" % (
                info["index"], info["comp"], self.status_name(info["status"]), info["frames"], info["expected"]))
        self.log("-----------------")

    def closeEvent(self, event):
        self.polling = False
        try:
            self._log_file.close()
        except Exception:
            pass
        super(RenderSpikeWidget, self).closeEvent(event)


def count_sequence_frames(path):
    """
        Number of files on disk for an image sequence output path such as
        "C:/renders/shot_[#####].exr". None for a movie (single file) path.
    """
    folder, name = os.path.split(path)
    match = SEQUENCE_TOKEN.search(name)
    if not match:
        return None
    pattern = re.compile("^%s\\d{%d,}%s$" % (re.escape(name[:match.start()]), len(match.group(1)),
                                             re.escape(name[match.end():])), re.IGNORECASE)
    try:
        return sum(1 for entry in os.listdir(folder) if pattern.match(entry))
    except OSError:
        return 0


def run(app=None):
    """Open the spike window. app: any Toolkit app bundle (defaults to the first one)."""
    engine = sgtk.platform.current_engine()
    if app is None:
        app = list(engine.apps.values())[0]
    return engine.show_dialog("Render Queue Spike", app, RenderSpikeWidget)
