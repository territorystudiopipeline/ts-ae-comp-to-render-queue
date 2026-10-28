# Copyright (c) 2026 Territory Studio
"""
Run ExtendScript in the open After Effects through ``AfterFX.com -ro``.

This module deliberately has no Qt and no ``engine.adobe`` dependency, so it is safe
to call from a worker thread. Anything that needs the After Effects RPC bridge
(the AE version, the project path) has to be read on the GUI thread first and
passed in.

How a run works (same approach as the Scene FX Manager):

* A small bootstrap ``.jsx`` is written to a temp folder. It puts the arguments on
  ``$.global.TS_JSX_ARGS``, runs the real script with ``$.evalFile`` and writes a
  JSON result file (``ok`` / ``error`` / ``warnings``).
* ``AfterFX.com -ro <bootstrap>`` is run with a timeout and no console window.
* Python reads the result file and raises :class:`AfterFXError` if the script
  failed, threw, never ran (usually a modal dialog open in After Effects) or timed out.
"""
import glob
import json
import os
import shutil
import subprocess
import tempfile
import threading
import time

try:
    import sgtk

    logger = sgtk.platform.get_logger(__name__)
except Exception:  # pragma: no cover - allows the module to be tested without Toolkit
    import logging

    logger = logging.getLogger(__name__)


DEFAULT_TIMEOUT = 300

# After Effects only runs one script at a time, so never send two at once.
_RUN_LOCK = threading.Lock()

# Major versions whose install folder is not "Adobe After Effects <2000 + major>".
_LEGACY_FOLDER_NAMES = {
    "16": "Adobe After Effects CC 2019",
    "17": "Adobe After Effects 2020",
    "18": "Adobe After Effects 2021",
}


class AfterFXError(RuntimeError):
    """Raised when a script could not be run in After Effects or reported a failure."""


def install_folder_name(ae_version):
    """
    The install folder name for an After Effects version string such as ``"25.2x13"``.

    :param str ae_version: ``app.version`` as reported by After Effects.
    :returns: e.g. ``"Adobe After Effects 2025"``, or None if the version can't be read.
    """
    major = str(ae_version or "").split(".")[0].split("x")[0].strip()
    if not major.isdigit():
        return None
    if major in _LEGACY_FOLDER_NAMES:
        return _LEGACY_FOLDER_NAMES[major]
    if int(major) >= 22:
        return "Adobe After Effects %d" % (2000 + int(major))
    return None


def find_afterfx(ae_version, override_path=None):
    """
    Find the ``AfterFX.com`` that matches the running After Effects.

    The matching version matters: ``-ro`` launches the executable it is given, so a
    different version would start a second After Effects instead of talking to the
    open one. If no exact match is found this raises rather than guessing.

    :param str ae_version: ``app.version`` of the running After Effects.
    :param str override_path: Explicit path from the ``afterfx_path`` app setting.
    :returns: Absolute path to ``AfterFX.com``.
    :raises AfterFXError: If it can't be found.
    """
    if override_path:
        override_path = os.path.expandvars(os.path.expanduser(override_path))
        if os.path.isfile(override_path):
            return override_path
        raise AfterFXError("The afterfx_path setting points to a missing file: %s" % override_path)

    if os.name != "nt":
        raise AfterFXError(
            "Running ExtendScript through AfterFX.com is only supported on Windows. "
            "Set the afterfx_path setting to override."
        )

    folder_name = install_folder_name(ae_version)
    if not folder_name:
        raise AfterFXError("Could not work out the After Effects install folder from version '%s'." % ae_version)

    roots = []
    for var in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        value = os.environ.get(var)
        if value and value not in roots:
            roots.append(value)
    if r"C:\Program Files" not in roots:
        roots.append(r"C:\Program Files")

    tried = []
    for root in roots:
        candidate = os.path.join(root, "Adobe", folder_name, "Support Files", "AfterFX.com")
        tried.append(candidate)
        if os.path.isfile(candidate):
            return candidate

    # Last resort: an install with a suffix, e.g. "Adobe After Effects 2025 (Beta)".
    for root in roots:
        pattern = os.path.join(root, "Adobe", folder_name + "*", "Support Files", "AfterFX.com")
        matches = sorted(glob.glob(pattern))
        if len(matches) == 1:
            return matches[0]

    raise AfterFXError("AfterFX.com not found. Looked in:\n  %s" % "\n  ".join(tried))


def _js_string(value):
    """A JavaScript string literal for ``value`` (JSON string literals are valid JS)."""
    return json.dumps(str(value), ensure_ascii=True)


def build_bootstrap(script_path, args, result_path):
    """
    The bootstrap ExtendScript that passes ``args`` to ``script_path`` and records the outcome.

    Kept free of ``JSON`` in ExtendScript so it works whether or not the host has it.
    """
    script_path = script_path.replace("\\", "/")
    result_path = result_path.replace("\\", "/")
    args_literal = json.dumps(args if args is not None else {}, ensure_ascii=True)

    return """// Generated by ts-ae-comp-to-render-queue. Safe to delete.
(function () {
    function esc(s) {
        return String(s).replace(/\\\\/g, "\\\\\\\\").replace(/"/g, '\\\\"')
                        .replace(/\\r/g, "\\\\r").replace(/\\n/g, "\\\\n");
    }
    function writeResult(ok, error) {
        var warnings = $.global.TS_JSX_WARNINGS || [];
        var parts = [];
        for (var i = 0; i < warnings.length; i++) {
            parts.push('"' + esc(warnings[i]) + '"');
        }
        var text = '{"ok": ' + (ok ? 'true' : 'false') +
                   ', "error": "' + esc(error || "") + '"' +
                   ', "warnings": [' + parts.join(", ") + ']}';
        var f = new File(%(result)s);
        f.encoding = "UTF-8";
        if (f.open("w")) {
            f.write(text);
            f.close();
        }
    }
    $.global.TS_JSX_ARGS = %(args)s;
    $.global.TS_JSX_WARNINGS = [];
    try {
        $.evalFile(new File(%(script)s));
        writeResult(true, "");
    } catch (e) {
        writeResult(false, e.toString() + (e.line ? " (line " + e.line + ")" : ""));
    } finally {
        $.global.TS_JSX_ARGS = undefined;
        $.global.TS_JSX_WARNINGS = undefined;
    }
})();
""" % {
        "result": _js_string(result_path),
        "args": args_literal,
        "script": _js_string(script_path),
    }


def run_script(afterfx_path, script_path, args=None, timeout=DEFAULT_TIMEOUT, strict=True, log=None):
    """
    Run ``script_path`` in the open After Effects and wait for it to finish.

    Thread safe (runs are serialised) and Qt free, so it can run in a worker thread.

    :param str afterfx_path: ``AfterFX.com`` from :func:`find_afterfx`.
    :param str script_path: The ``.jsx`` to run.
    :param args: JSON-serialisable value exposed to the script as ``$.global.TS_JSX_ARGS``.
    :param int timeout: Seconds to wait for After Effects.
    :param bool strict: Treat warnings the script recorded as a failure.
    :param log: Logger-like object to use. In a worker thread pass one that hands records
        to the GUI thread: in tk-aftereffects every Toolkit log record is sent to After
        Effects over the RPC socket, which must only be used from the GUI thread.
    :returns: The list of warnings the script recorded.
    :raises AfterFXError: On any failure.
    """
    log = log or logger
    if not os.path.isfile(script_path):
        raise AfterFXError("JSX script not found: %s" % script_path)
    timeout = timeout or DEFAULT_TIMEOUT

    temp_dir = tempfile.mkdtemp(prefix="ts_ae_rq_")
    bootstrap_path = os.path.join(temp_dir, "bootstrap.jsx")
    result_path = os.path.join(temp_dir, "result.json")
    keep_temp = False

    try:
        with open(bootstrap_path, "w", encoding="utf-8") as handle:
            handle.write(build_bootstrap(script_path, args, result_path))

        command = [afterfx_path, "-ro", bootstrap_path]
        kwargs = {"capture_output": True, "text": True, "timeout": timeout}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        log.debug("Running in After Effects: %s (args: %s)", script_path, args)
        start = time.time()
        with _RUN_LOCK:
            try:
                completed = subprocess.run(command, **kwargs)
            except subprocess.TimeoutExpired:
                keep_temp = True
                raise AfterFXError(
                    "After Effects did not finish %s within %s seconds. Close any open dialog "
                    "in After Effects and try again." % (os.path.basename(script_path), timeout)
                )
            except OSError as exc:
                raise AfterFXError("Could not start %s: %s" % (afterfx_path, exc))

            # AfterFX.com normally waits for the script, but give the result file time to
            # appear in case it hands over and returns early. A non-zero exit means the
            # script was most likely never delivered, so only wait briefly then.
            wait_until = start + timeout if completed.returncode == 0 else time.time() + 2
            while not os.path.isfile(result_path) and time.time() < wait_until:
                time.sleep(0.25)

        if completed.returncode != 0:
            log.warning(
                "AfterFX.com exited with code %s\nstdout: %s\nstderr: %s",
                completed.returncode, completed.stdout, completed.stderr,
            )

        if not os.path.isfile(result_path):
            keep_temp = True
            raise AfterFXError(
                "After Effects did not run %s (exit code %s). Close any open dialog in After "
                "Effects and try again.\n%s" % (
                    os.path.basename(script_path), completed.returncode,
                    (completed.stderr or completed.stdout or "").strip(),
                )
            )

        with open(result_path, "r", encoding="utf-8-sig") as handle:
            result = json.load(handle)

        warnings = result.get("warnings") or []
        for warning in warnings:
            log.warning("%s: %s", os.path.basename(script_path), warning)

        if not result.get("ok"):
            raise AfterFXError("%s failed: %s" % (os.path.basename(script_path), result.get("error")))
        if strict and warnings:
            raise AfterFXError("%s reported problems:\n%s" % (os.path.basename(script_path), "\n".join(warnings)))

        log.debug("%s finished in %.1fs", os.path.basename(script_path), time.time() - start)
        return warnings

    finally:
        if keep_temp:
            log.debug("Kept After Effects run files for debugging: %s", temp_dir)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
