"""
Render output paths, replacing the Toolkit templates (``ae_<entity>_work_*_render`` and
``ae_<entity>_work_*render_scene`` in ts-config-default2).

In Toolkit every render template ended the same way and only the render root changed with the
entity type (``@shot_aftereffects_render_root`` and so on). Here that becomes:

- ``render_root_profiles``: the render root for the current AYON folder, picked by folder type
  (first match wins, an empty ``folder_types`` list matches every type).
- ``project_render_root``: the render root when the context has no folder (Toolkit's project
  environment, the ``artists/{User}`` area).
- six path templates that use ``{render_root}``.

Templates are Python format strings filled with AYON template data (``{root[work]}``,
``{project[name]}``, ``{folder[path]}``, ``{user[name]}`` ...) plus ``name``, ``version``, ``ext``,
``ae_comp_name`` and ``frame``. ``name`` and ``version`` come from the work file name, as before.

Only :func:`get_ayon_template_data` needs AYON; the rest is plain Python so it can be tested alone.
"""
import os
import re
import string


class OutputPathError(Exception):
    """An output path can't be built (unexpected work file name, missing template key ...)."""


# (is movie, use comp name, render scene) -> template setting
TEMPLATE_KEYS = {
    (True, True, True): "comp_render_scene",
    (True, True, False): "mov_render_comp",
    (True, False, True): "render_scene",
    (True, False, False): "mov_render",
    (False, True, True): "comp_render_scene",
    (False, True, False): "seq_render_comp",
    (False, False, True): "render_scene",
    (False, False, False): "seq_render",
}


def preset_extension(preset_path):
    """The output extension a preset file stands for: ``mov_prores422_RGB.aep`` -> ``mov``."""
    return os.path.basename(preset_path).split("_")[0]


def is_movie_preset(preset_path, movie_formats):
    file_name = os.path.basename(preset_path)
    return any(file_name.startswith(fmt) for fmt in movie_formats)


def template_key(preset_path, movie_formats, use_comp_name, render_scene):
    """Which of the six path templates applies."""
    return TEMPLATE_KEYS[(is_movie_preset(preset_path, movie_formats), bool(use_comp_name), bool(render_scene))]


def parse_workfile_name(file_name, regex):
    """
    ``(name, version)`` from the work file name, e.g. ``sh010_comp_v003.aep`` -> ``("comp", 3)``.

    :param regex: Pattern with ``name`` and ``version`` groups (``workfile_regex`` setting).
    """
    match = re.match(regex, file_name or "")
    if not match:
        raise OutputPathError(
            "Couldn't read the name and version from the project file name '%s'. "
            "Save the project as a work file first." % file_name)
    return match.group("name"), int(match.group("version"))


def _format(template, data, setting_name):
    try:
        return string.Formatter().vformat(template, (), data)
    except KeyError as error:
        raise OutputPathError(
            "The '%s' template needs %s, which the current context doesn't have." % (setting_name, error))
    except (IndexError, ValueError, TypeError) as error:
        raise OutputPathError("The '%s' template can't be filled: %s" % (setting_name, error))


def render_root(templates, template_data):
    """The render root for the context in ``template_data``."""
    folder = template_data.get("folder")
    if not folder:
        return _format(templates["project_render_root"], template_data, "project_render_root")

    folder_type = (folder.get("type") or "").lower()
    for profile in templates.get("render_root_profiles", []):
        types = [value.lower() for value in profile.get("folder_types", [])]
        if not types or folder_type in types:
            return _format(profile["render_root"], template_data, "render_root_profiles")
    raise OutputPathError("No render root is set for folder type '%s'." % folder.get("type"))


def build_output_path(templates, template_data, preset_path, movie_formats, workfile_name,
                      comp_name, use_comp_name, render_scene):
    """
    The render output path (or render scene ``.aep`` path) for one render queue item.

    :param dict templates: The ``output_templates`` setting.
    :param dict template_data: AYON template data for the current context (see get_ayon_template_data).
    :param str preset_path: The output module preset ``.aep``; its file name gives the extension.
    :param list movie_formats: ``render_preset_movie_formats``; other presets render image sequences.
    :param str workfile_name: The After Effects project file name.
    :param str comp_name: The comp's name, used when ``use_comp_name`` is set.
    :param bool render_scene: Return the render scene ``.aep`` path instead of the output path.
    """
    key = template_key(preset_path, movie_formats, use_comp_name, render_scene)
    name, version = parse_workfile_name(workfile_name, templates["workfile_regex"])

    data = dict(template_data)
    data.update({
        "name": name,
        "version": version,
        "ext": preset_extension(preset_path),
        "ae_comp_name": comp_name,
        "frame": templates.get("sequence_frame_token", "[####]"),
    })
    data["render_root"] = render_root(templates, data)
    return os.path.normpath(_format(templates[key], data, key))


def get_ayon_template_data():
    """
    AYON template data for the current context, with ``root`` added (as plain strings for
    this platform). Only works inside an AYON host session.
    """
    from ayon_core.pipeline import Anatomy, get_current_project_name
    from ayon_core.pipeline.context_tools import get_current_context_template_data

    data = get_current_context_template_data()
    anatomy = Anatomy(get_current_project_name())
    data["root"] = {name: str(root) for name, root in anatomy.roots.items()}
    return data
