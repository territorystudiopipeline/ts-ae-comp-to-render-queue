"""
Settings for the render queue tool, replacing Toolkit's ``app.get_setting``.

Defaults ship next to this module in ``settings_defaults.json`` (taken from the
``ts-ae-comp-to-render-queue`` settings in ts-config-default2). Extra JSON files listed in the
``TS_COMP_TO_RQ_SETTINGS`` environment variable (``os.pathsep`` separated, later files win) are
merged on top: dictionaries merge key by key, any other value (lists included) is replaced.

Paths may start with ``{package_root}``, the root of this package (it replaces Toolkit's ``{self}``).

This module only uses the standard library, so it can be imported and tested outside AYON.
"""
import copy
import json
import logging
import os

logger = logging.getLogger(__name__)

ENV_VAR = "TS_COMP_TO_RQ_SETTINGS"
PACKAGE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEFAULTS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings_defaults.json")


class SettingsError(Exception):
    """A settings file is missing or isn't valid JSON."""


def deep_merge(base, override):
    """
    Return a copy of ``base`` with ``override`` merged in. Dictionaries merge key by key;
    anything else in ``override`` replaces the value in ``base``.
    """
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def resolve_package_path(path):
    """Expand ``{package_root}`` in a path and normalise it for this platform."""
    return os.path.normpath(path.replace("{package_root}", PACKAGE_ROOT))


def _read_json(path):
    if not os.path.isfile(path):
        raise SettingsError("Settings file not found: %s" % path)
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except ValueError as error:
        raise SettingsError("Settings file %s isn't valid JSON: %s" % (path, error))
    if not isinstance(data, dict):
        raise SettingsError("Settings file %s must hold a JSON object" % path)
    return data


def override_paths_from_env(environ=None):
    """The override files named in ``TS_COMP_TO_RQ_SETTINGS``, in order."""
    value = (environ if environ is not None else os.environ).get(ENV_VAR, "")
    return [path for path in value.split(os.pathsep) if path.strip()]


class Settings(object):
    """Read-only access to the merged settings."""

    def __init__(self, data, sources=()):
        self._data = data
        self.sources = list(sources)

    def get(self, key, default=None):
        """A copy of a setting, so callers can't change the loaded settings by accident."""
        return copy.deepcopy(self._data.get(key, default))

    # Same call shape as Toolkit, so the dialog can switch with a one-line change.
    get_setting = get

    def render_presets(self):
        """``[(name, absolute .aep path), ...]`` in settings order."""
        return [(preset["name"], resolve_package_path(preset["path"])) for preset in self._data.get("render_presets", [])]

    def as_dict(self):
        return copy.deepcopy(self._data)


def load_settings(override_paths=None):
    """
    Load the defaults and merge the override files on top.

    :param override_paths: Override JSON files, later ones win. None reads ``TS_COMP_TO_RQ_SETTINGS``.
    :raises SettingsError: An override file is missing or isn't valid JSON. A wrong path fails
        loudly rather than silently falling back to the defaults (e.g. the wrong Deadline host).
    """
    if override_paths is None:
        override_paths = override_paths_from_env()

    data = _read_json(DEFAULTS_PATH)
    sources = [DEFAULTS_PATH]
    for path in override_paths:
        data = deep_merge(data, _read_json(path))
        sources.append(path)
    logger.debug("Settings loaded from: %s" % ", ".join(sources))
    return Settings(data, sources)
