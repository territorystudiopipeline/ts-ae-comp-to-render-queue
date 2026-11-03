# Copyright (c) 2026 Territory Studio
"""
SPIKE (temporary): can the farm swap comp proxies in, with OCIO-correct read-back?

Three steps, each writing JSON and renders into a farm-visible spike folder, and copying the
results into the Toolkit log folder (ts_proxy_spike/) so they can be reviewed:

1. Inspect: what scripting can see / set in this After Effects: project colour (OCIO)
   properties, footage source properties, output module and render settings keys, Proxy Use.
2. Local proxy test (in a NEW, EMPTY project): renders a small precomp to EXR and PNG, sets each
   as the comp proxy (conform fps, premultiplied alpha), renders hosts that use the precomp
   (normal, collapsed, 3D) with "Use Comp Proxies Only", plus a live reference. Saves the scratch
   project as proxy_spike.aep for step 3.
3. Farm patch test: submits a Deadline CommandLine job running `AfterFX.exe -noui -r patch.jsx`
   (opens the scene copy, sets the proxy, saves a patched copy), then a dependent After Effects
   job rendering the host from the patched copy. Polls both with DeadlineConnect GetJobs and
   logs the raw job fields (checks the job monitor's assumptions too).

Run: set TS_AE_PROXY_SPIKE=1 before launching After Effects -> "Proxy Spike..." in the
ShotGrid menu.
"""
import datetime
import json
import os
import shutil
import tempfile
import time
import traceback

import sgtk
from sgtk.platform.qt import QtCore, QtGui

from . import afterfx

JSX_DIR = os.path.join(os.path.dirname(__file__), "spike_proxy_jsx")
FPS = 23.976
FRAMES = 10
SIZE = 256
TEST_FRAME = 5


def _log_folder():
    try:
        return sgtk.LogManager().log_folder
    except Exception:
        return tempfile.gettempdir()


def _combined_jsx(name):
    """A script with common.jsxinc inlined (no #include, so it runs from anywhere)."""
    with open(os.path.join(JSX_DIR, "common.jsxinc"), encoding="utf-8") as handle:
        common = handle.read()
    with open(os.path.join(JSX_DIR, name), encoding="utf-8") as handle:
        body = handle.read()
    lines = [line for line in body.splitlines() if not line.startswith("#include")]
    return common + "\n" + "\n".join(lines) + "\n"


class _Task(QtCore.QObject):
    """Runs a function on a worker thread; it reports through log lines only."""
    line = QtCore.Signal(str)
    done = QtCore.Signal(object)

    def __init__(self, func):
        super(_Task, self).__init__()
        self.func = func

    @QtCore.Slot()
    def run(self):
        try:
            result = self.func(self.line.emit)
        except Exception:
            self.line.emit("FAILED:\n" + traceback.format_exc())
            result = None
        self.done.emit(result)


_THREADS = set()  # threads kept alive until they finish, even if the window closes


class ProxySpikeWidget(QtGui.QWidget):

    def __init__(self, parent=None):
        super(ProxySpikeWidget, self).__init__(parent)
        self._app = sgtk.platform.current_bundle()
        self.engine = self._app.engine
        self.adobe = self.engine.adobe
        self.settings = QtCore.QSettings("Territory", "AEProxySpike")
        self.results_dir = os.path.join(_log_folder(), "ts_proxy_spike")
        os.makedirs(self.results_dir, exist_ok=True)
        self.log_path = os.path.join(self.results_dir, "spike_%s.log" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        self._busy = False
        self._build_ui()
        self.log("Results are copied to %s" % self.results_dir)
        self.log("After Effects %s" % self.adobe.app.version)

    # ------------------------------------------------------------------ ui

    def _build_ui(self):
        layout = QtGui.QVBoxLayout(self)
        form = QtGui.QFormLayout()
        self.folder_edit = QtGui.QLineEdit(self.settings.value("folder", ""))
        self.folder_edit.setPlaceholderText(r"A folder the farm can reach, e.g. \\ldn-fs5\projects-a\...\renders\_proxy_spike")
        form.addRow("Spike folder (farm visible)", self.folder_edit)
        self.pool_edit = QtGui.QLineEdit(self.settings.value("pool", self._default_pool()))
        self.group_edit = QtGui.QLineEdit(self.settings.value("group", (self._app.get_setting("deadline_defaults") or {}).get("group", "ae")))
        form.addRow("Deadline pool", self.pool_edit)
        form.addRow("Deadline group", self.group_edit)
        self.afterfx_edit = QtGui.QLineEdit(self.settings.value("afterfx_exe", self._default_afterfx_exe()))
        form.addRow("AfterFX.exe on the farm", self.afterfx_edit)
        self.colour_edit = QtGui.QLineEdit(self.settings.value("colour_tries", "{}"))
        self.colour_edit.setToolTip('Optional proxy colour properties to try in steps 2 and 3, as JSON:\n'
                                    '{"exr": {"propertyName": value}, "png": {...}} (filled in after step 1)')
        form.addRow("Colour settings to try", self.colour_edit)
        layout.addLayout(form)

        row = QtGui.QHBoxLayout()
        self.inspect_button = QtGui.QPushButton("1. Inspect scripting")
        self.local_button = QtGui.QPushButton("2. Local proxy test (new empty project)")
        self.farm_button = QtGui.QPushButton("3. Farm patch test")
        for button in (self.inspect_button, self.local_button, self.farm_button):
            row.addWidget(button)
        layout.addLayout(row)

        self.output = QtGui.QPlainTextEdit()
        self.output.setReadOnly(True)
        self.output.setLineWrapMode(QtGui.QPlainTextEdit.NoWrap)
        font = QtGui.QFont("Consolas")
        font.setStyleHint(QtGui.QFont.Monospace)
        self.output.setFont(font)
        layout.addWidget(self.output)
        self.resize(980, 560)

        self.inspect_button.clicked.connect(self.run_inspect)
        self.local_button.clicked.connect(self.run_local_test)
        self.farm_button.clicked.connect(self.run_farm_test)

    def _default_pool(self):
        try:
            data = self._app.shotgun.find_one("Project", [["id", "is", self._app.context.project["id"]]], ["sg_ae_render_pool"])
            return (data or {}).get("sg_ae_render_pool") or "none"
        except Exception:
            return "none"

    def _default_afterfx_exe(self):
        try:
            com = afterfx.find_afterfx(self.adobe.app.version)
            return os.path.join(os.path.dirname(com), "AfterFX.exe")
        except Exception:
            return ""

    def log(self, text):
        line = "%s  %s" % (datetime.datetime.now().strftime("%H:%M:%S"), text)
        with open(self.log_path, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
        self.output.appendPlainText(line)

    def _save_fields(self):
        for key, widget in (("folder", self.folder_edit), ("pool", self.pool_edit), ("group", self.group_edit),
                            ("afterfx_exe", self.afterfx_edit), ("colour_tries", self.colour_edit)):
            self.settings.setValue(key, widget.text().strip())

    def _folder(self):
        folder = self.folder_edit.text().strip().rstrip("\\/")
        if not folder:
            raise ValueError("Set the spike folder first (a folder the farm can reach).")
        os.makedirs(folder, exist_ok=True)
        return folder

    def _colour_tries(self):
        text = self.colour_edit.text().strip() or "{}"
        return json.loads(text)

    def _start(self, func, on_done=None):
        if self._busy:
            self.log("Still running the previous step.")
            return
        self._save_fields()
        self._busy = True
        for button in (self.inspect_button, self.local_button, self.farm_button):
            button.setEnabled(False)
        thread = QtCore.QThread()
        task = _Task(func)
        task.moveToThread(thread)
        thread.started.connect(task.run)
        task.line.connect(self.log)
        task.done.connect(lambda result: self._finished(result, on_done))
        task.done.connect(thread.quit, QtCore.Qt.DirectConnection)
        thread.finished.connect(lambda t=thread: _THREADS.discard(t))
        _THREADS.add(thread)
        self._task = task
        thread.start()

    def _finished(self, result, on_done):
        self._busy = False
        for button in (self.inspect_button, self.local_button, self.farm_button):
            button.setEnabled(True)
        if on_done and result is not None:
            on_done(result)

    def _copy_results(self, paths, emit):
        for path in paths:
            if os.path.isfile(path):
                target = os.path.join(self.results_dir, os.path.basename(path))
                shutil.copy2(path, target)
                emit("copied %s" % os.path.basename(path))
            else:
                emit("missing %s" % path)

    def _run_jsx(self, name, args, emit, timeout=600):
        afterfx_com = afterfx.find_afterfx(self.adobe.app.version)
        temp_dir = tempfile.mkdtemp(prefix="ts_proxy_spike_")
        script = os.path.join(temp_dir, name)
        with open(script, "w", encoding="utf-8") as handle:
            handle.write(_combined_jsx(name))
        emit("Running %s in After Effects..." % name)
        start = time.time()
        warnings = afterfx.run_script(afterfx_com, script, args=args, timeout=timeout, strict=False, log=_EmitLog(emit))
        emit("%s finished in %.1fs%s" % (name, time.time() - start, (" (warnings: %s)" % warnings) if warnings else ""))

    # ------------------------------------------------------------------ step 1

    def run_inspect(self):
        try:
            folder = self._folder()
        except Exception as e:
            self.log(str(e))
            return
        out = os.path.join(folder, "inspect.json")

        def work(emit):
            self._run_jsx("inspect.jsx", {"out": out}, emit)
            data = _load(out)
            emit("Project colour values: %s" % json.dumps(data.get("projectValues"), indent=1))
            emit("Proxy Use set -> %s" % [s for s in data.get("steps", []) if "Proxy Use" in s["name"]])
            failed = [s for s in data.get("steps", []) if not s.get("ok")]
            if failed:
                emit("Failed steps: %s" % json.dumps(failed, indent=1))
            self._copy_results([out], emit)
            return True

        self._start(work)

    # ------------------------------------------------------------------ step 2

    def run_local_test(self):
        try:
            folder = self._folder()
            colour_tries = self._colour_tries()
        except Exception as e:
            self.log(str(e))
            return
        if self.adobe.app.project.numItems > 0:
            self.log("Step 2 needs a NEW, EMPTY project: File > New > New Project, then run it again.")
            return
        out = os.path.join(folder, "local_result.json")
        args = {"out": out, "folder": folder.replace("\\", "/"), "fps": FPS, "frames": FRAMES, "size": SIZE,
                "testFrame": TEST_FRAME, "colourTries": colour_tries}

        def work(emit):
            self._run_jsx("local_test.jsx", args, emit, timeout=900)
            data = _load(out)
            for step in data.get("steps", []):
                emit("%s %s: %s" % ("ok  " if step.get("ok") else "FAIL", step["name"],
                                    json.dumps(step.get("value") if step.get("ok") else step.get("error"))[:300]))
            verify = os.path.join(folder, "verify")
            files = [os.path.join(verify, name) for name in sorted(os.listdir(verify))] if os.path.isdir(verify) else []
            first_proxy = [os.path.join(folder, "proxies", fmt, sorted(os.listdir(os.path.join(folder, "proxies", fmt)))[0])
                           for fmt in ("exr", "png") if os.listdir(os.path.join(folder, "proxies", fmt))]
            self._copy_results([out] + files + first_proxy, emit)
            emit("Step 2 done. proxy_spike.aep is saved in the spike folder for step 3.")
            return True

        self._start(work)

    # ------------------------------------------------------------------ step 3

    def run_farm_test(self):
        try:
            folder = self._folder()
            colour_tries = self._colour_tries().get("exr", {})
        except Exception as e:
            self.log(str(e))
            return
        scene = os.path.join(folder, "proxy_spike.aep")
        exr_dir = os.path.join(folder, "proxies", "exr")
        if not os.path.isfile(scene) or not os.path.isdir(exr_dir) or not os.listdir(exr_dir):
            self.log("Run step 2 first (needs proxy_spike.aep and the EXR proxy frames in the spike folder).")
            return
        first_frame = os.path.join(exr_dir, sorted(os.listdir(exr_dir))[0])
        afterfx_exe = self.afterfx_edit.text().strip()
        pool, group = self.pool_edit.text().strip(), self.group_edit.text().strip()
        host, port = self._app.get_setting("deadline_host"), self._app.get_setting("deadline_port")
        version = ".".join(self.adobe.app.version.split("x")[0].split(".")[:2])
        stamp = datetime.datetime.now().strftime("%H%M%S")
        patched = os.path.join(folder, "farm", "proxy_spike_patched_%s.aep" % stamp)
        patch_result = os.path.join(folder, "farm", "patch_result_%s.json" % stamp)
        patch_jsx = os.path.join(folder, "farm", "patch_%s.jsx" % stamp)
        os.makedirs(os.path.dirname(patch_jsx), exist_ok=True)

        config = {"scene": scene.replace("\\", "/"), "patched": patched.replace("\\", "/"),
                  "result": patch_result.replace("\\", "/"), "firstFrame": first_frame.replace("\\", "/"),
                  "precomp": "P_precomp", "host": "H_normal", "colourTries": colour_tries}
        with open(os.path.join(JSX_DIR, "patch_template.jsx"), encoding="utf-8") as handle:
            template = handle.read()
        with open(os.path.join(JSX_DIR, "common.jsxinc"), encoding="utf-8") as handle:
            common = handle.read()
        script = template.replace('#include "%(common)s"', common).replace("%(config)s", json.dumps(config))
        with open(patch_jsx, "w", encoding="utf-8") as handle:
            handle.write(script)
        self.log("Wrote %s" % patch_jsx)

        self_version = self.adobe.app.version

        def work(emit):
            import Deadline.DeadlineConnect as Connect
            deadline = Connect.DeadlineCon(host, port)
            common_job = {"Pool": pool, "Group": group, "UserName": os.environ.get("USERNAME", ""),
                          "BatchName": "TS proxy spike %s" % stamp, "TaskTimeoutMinutes": 20, "MachineLimit": 1}
            patch_job = dict(common_job, Plugin="CommandLine", Name="proxy spike %s - patch (AfterFX -noui)" % stamp,
                             Frames="0")
            patch_plugin = {"Executable": afterfx_exe, "Arguments": '-noui -r "%s"' % patch_jsx,
                            "Shell": "default", "ShellExecute": False, "StartupDirectory": ""}
            emit("Submitting patch job: %s %s" % (afterfx_exe, patch_plugin["Arguments"]))
            job1 = deadline.Jobs.SubmitJob(patch_job, patch_plugin)
            if not isinstance(job1, dict) or "_id" not in job1:
                raise RuntimeError("Patch job refused: %s" % job1)
            render_job = dict(common_job, Plugin="AfterEffects", Name="proxy spike %s - render H_normal (patched)" % stamp,
                              Frames=str(TEST_FRAME), ChunkSize=1, JobDependencies=job1["_id"])
            render_plugin = {"SceneFile": patched, "Comp": "H_normal", "Version": version,
                             "SubmittedFromVersion": self_version}
            job2 = deadline.Jobs.SubmitJob(render_job, render_plugin)
            if not isinstance(job2, dict) or "_id" not in job2:
                raise RuntimeError("Render job refused: %s" % job2)
            ids = [job1["_id"], job2["_id"]]
            emit("Submitted patch %s and render %s; polling every 10s (up to 30 min)" % tuple(ids))

            logged_raw = False
            deadline_end = time.time() + 30 * 60
            last = {}
            while time.time() < deadline_end:
                start = time.time()
                jobs = deadline.Jobs.GetJobs(ids)
                elapsed = time.time() - start
                if not isinstance(jobs, list):
                    emit("GetJobs returned %r" % (jobs,))
                    time.sleep(10)
                    continue
                if not logged_raw and jobs:
                    logged_raw = True
                    emit("GetJobs(list of 2) took %.2fs and returned %d jobs" % (elapsed, len(jobs)))
                    sample = jobs[0]
                    emit("Job keys: %s" % sorted(sample.keys()))
                    emit("Stat=%r CompletedChunks=%r RenderingChunks=%r QueuedChunks=%r FailedChunks=%r Props.Tasks=%r" % (
                        sample.get("Stat"), sample.get("CompletedChunks"), sample.get("RenderingChunks"),
                        sample.get("QueuedChunks"), sample.get("FailedChunks"), (sample.get("Props") or {}).get("Tasks")))
                states = {}
                for job in jobs:
                    states[job["_id"]] = (job.get("Stat"), job.get("CompletedChunks"), job.get("FailedChunks"))
                if states != last:
                    emit("states: %s" % states)
                    last = states
                stats = [s[0] for s in states.values()]
                if len(stats) == 2 and all(s in (3, 4) for s in stats):
                    break
                if any(s == 4 for s in stats) or any((s[2] or 0) > 0 for s in states.values()):
                    emit("A job failed or has failed tasks: check its task report in the Monitor.")
                    if states.get(ids[0], (None,))[0] == 4 or (states.get(ids[0], (0, 0, 0))[2] or 0) >= 1:
                        break
                time.sleep(10)

            farm_dir = os.path.join(folder, "farm")
            outputs = [os.path.join(farm_dir, name) for name in os.listdir(farm_dir)
                       if name.startswith("farm_H_normal")]
            self._copy_results([patch_result] + outputs, emit)
            if os.path.isfile(patch_result):
                emit("Patch result: %s" % json.dumps(_load(patch_result), indent=1)[:3000])
            return True

        self._start(work)


class _EmitLog(object):
    """Logger-like object for afterfx.run_script on a worker thread."""

    def __init__(self, emit):
        self.emit = emit

    def _log(self, message, *args):
        self.emit(message % args if args else message)

    debug = info = warning = error = _log


def _load(path):
    with open(path, encoding="utf-8-sig") as handle:
        return json.load(handle)


def run(app=None):
    engine = sgtk.platform.current_engine()
    if app is None:
        app = list(engine.apps.values())[0]
    return engine.show_dialog("Proxy Spike", app, ProxySpikeWidget)
