"""
Tests for python/app/settings.py and python/app/output_paths.py.

The modules are loaded by file path: importing the ``app`` package would pull in the Toolkit
dialog. Run with:  python -m unittest discover tests
"""
import importlib.util
import json
import os
import tempfile
import unittest

APP_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "python", "app")


def _load(name):
    spec = importlib.util.spec_from_file_location("ts_rq_" + name, os.path.join(APP_DIR, name + ".py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


settings_mod = _load("settings")
output_paths = _load("output_paths")

ROOT = "/mnt/projects-a"


def context(folder_path=None, folder_type=None):
    """AYON-style template data, shaped like ayon_core's get_template_data."""
    data = {
        "root": {"work": ROOT},
        "project": {"name": "ABC_project", "code": "abc"},
        "user": {"name": "jdoe"},
    }
    if folder_path:
        parents = folder_path.strip("/").split("/")
        data["folder"] = {"name": parents.pop(), "type": folder_type, "path": folder_path, "parents": parents}
        data["hierarchy"] = "/".join(parents)
    return data


class SettingsTests(unittest.TestCase):
    def test_defaults_match_the_toolkit_config(self):
        settings = settings_mod.load_settings([])
        self.assertEqual(settings.get("default_first_frame"), 1001)
        self.assertEqual(settings.get("deadline_host"), "ldn-dl-rcs")
        self.assertEqual(settings.get("deadline_port"), 8081)
        self.assertEqual(settings.get("deadline_defaults")["frames_per_task"], 35)
        self.assertEqual(settings.get_setting("missing", "fallback"), "fallback")

    def test_every_preset_file_exists(self):
        presets = settings_mod.load_settings([]).render_presets()
        self.assertEqual([name for name, _ in presets][0], "EXR")
        for name, path in presets:
            self.assertTrue(os.path.isfile(path), "%s: %s" % (name, path))

    def test_overrides_merge_dicts_and_replace_lists(self):
        with tempfile.TemporaryDirectory() as tmp:
            first = os.path.join(tmp, "studio.json")
            second = os.path.join(tmp, "project.json")
            with open(first, "w") as handle:
                json.dump({"deadline_defaults": {"priority": 50}, "qsettings_ignore_list": []}, handle)
            with open(second, "w") as handle:
                json.dump({"deadline_defaults": {"priority": 70}}, handle)
            env = {settings_mod.ENV_VAR: os.pathsep.join([first, second])}
            settings = settings_mod.load_settings(settings_mod.override_paths_from_env(env))

        defaults = settings.get("deadline_defaults")
        self.assertEqual(defaults["priority"], 70)
        self.assertEqual(defaults["group"], "ae")  # untouched keys survive
        self.assertEqual(settings.get("qsettings_ignore_list"), [])

    def test_missing_or_broken_override_fails_loudly(self):
        with self.assertRaises(settings_mod.SettingsError):
            settings_mod.load_settings(["/no/such/settings.json"])
        with tempfile.TemporaryDirectory() as tmp:
            broken = os.path.join(tmp, "broken.json")
            with open(broken, "w") as handle:
                handle.write("{not json")
            with self.assertRaises(settings_mod.SettingsError):
                settings_mod.load_settings([broken])

    def test_get_returns_a_copy(self):
        settings = settings_mod.load_settings([])
        settings.get("deadline_defaults")["priority"] = 99
        self.assertEqual(settings.get("deadline_defaults")["priority"], 30)


class OutputPathTests(unittest.TestCase):
    def setUp(self):
        settings = settings_mod.load_settings([])
        self.templates = settings.get("output_templates")
        self.movie_formats = settings.get("render_preset_movie_formats")

    def build(self, data, preset, use_comp_name=False, render_scene=False, workfile="sh010_comp_v003.aep"):
        return output_paths.build_output_path(
            self.templates, data, preset, self.movie_formats, workfile, "MainComp", use_comp_name, render_scene)

    def test_shot_paths_match_the_toolkit_templates(self):
        # ae_shot_work_*: @shot_root/_work/aftereffects/renders/{name}/v{version}/...
        renders = ROOT + "/ABC_project/sequences/sq010/sh010/_work/aftereffects/renders/comp/v003"
        data = context("/sequences/sq010/sh010", "Shot")
        exr, mov = "exr_workingSpace_RGBA.aep", "mov_prores422_RGB.aep"
        expected = {
            (exr, False, False): renders + "/exr/comp_v003.[####].exr",
            (mov, False, False): renders + "/mov/comp_v003.mov",
            (exr, False, True): renders + "/exr/aep/comp_v003.aep",
            (exr, True, False): renders + "/MainComp/exr/MainComp_v003.[####].exr",
            (mov, True, False): renders + "/MainComp/mov/MainComp_v003.mov",
            (mov, True, True): renders + "/MainComp/mov/aep/MainComp_v003.aep",
        }
        for (preset, use_comp_name, render_scene), path in expected.items():
            self.assertEqual(self.build(data, preset, use_comp_name, render_scene), os.path.normpath(path))

    def test_other_folder_types_use_their_folder_path(self):
        for folder_path, folder_type in (("/sequences/sq010", "Sequence"),
                                         ("/assets/characters/hero/heroA", "Asset")):
            path = self.build(context(folder_path, folder_type), "png_16bit_RGBA.aep")
            self.assertEqual(path, os.path.normpath(
                ROOT + "/ABC_project" + folder_path + "/_work/aftereffects/renders/comp/v003/png/comp_v003.[####].png"))

    def test_project_context_uses_the_artist_area(self):
        # ae_project_work_*: artists/{User}/aftereffects/renders
        path = self.build(context(), "mov_prores4444_RGB.aep", workfile="ABC_jd_layout_v012.aep")
        self.assertEqual(path, os.path.normpath(
            ROOT + "/ABC_project/artists/jdoe/aftereffects/renders/layout/v012/mov/layout_v012.mov"))

    def test_profiles_pick_by_folder_type(self):
        self.templates["render_root_profiles"] = [
            {"folder_types": ["shot"], "render_root": "{root[work]}/shots/{folder[name]}"},
            {"folder_types": [], "render_root": "{root[work]}/other"},
        ]
        self.assertTrue(self.build(context("/sq/sh010", "Shot"), "mov_x.aep").startswith(
            os.path.normpath(ROOT + "/shots/sh010/")))
        self.assertTrue(self.build(context("/assets/a", "Asset"), "mov_x.aep").startswith(
            os.path.normpath(ROOT + "/other/")))

        self.templates["render_root_profiles"] = [{"folder_types": ["Shot"], "render_root": "x"}]
        with self.assertRaises(output_paths.OutputPathError):
            self.build(context("/assets/a", "Asset"), "mov_x.aep")

    def test_unsaved_or_unexpected_work_file_names_are_reported(self):
        for name in ("Untitled Project.aep", "", None):
            with self.assertRaises(output_paths.OutputPathError):
                self.build(context("/sq/sh010", "Shot"), "mov_x.aep", workfile=name)

    def test_missing_template_key_is_reported(self):
        data = context("/sq/sh010", "Shot")
        del data["root"]
        with self.assertRaises(output_paths.OutputPathError) as caught:
            self.build(data, "mov_x.aep")
        self.assertIn("root", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
