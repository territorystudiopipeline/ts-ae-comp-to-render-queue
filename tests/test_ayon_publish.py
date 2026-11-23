"""
Tests for ayon_plugins/publish/collect_ts_render_frame_range.py and python/app/ayon_publish.py.

AYON and pyblish aren't needed: pyblish.api is replaced by a minimal stand-in when it isn't
installed, and After Effects by fake stubs. Run with:  python -m unittest discover tests
"""
import importlib.util
import json
import os
import sys
import types
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, os.path.join(REPO, path))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


try:
    import pyblish.api  # noqa: F401
except ImportError:
    api = types.ModuleType("pyblish.api")
    api.CollectorOrder = 0
    api.ContextPlugin = type("ContextPlugin", (object,), {"log": None})
    sys.modules["pyblish"] = types.ModuleType("pyblish")
    sys.modules["pyblish"].api = api
    sys.modules["pyblish.api"] = api

plugin = _load("ts_rq_collect_frame_range", "ayon_plugins/publish/collect_ts_render_frame_range.py")
ayon_publish = _load("ts_rq_ayon_publish", "python/app/ayon_publish.py")


class FakeInstance(object):
    def __init__(self, **data):
        self.data = data


class FakeLog(object):
    def __init__(self):
        self.messages = []

    def _log(self, message):
        self.messages.append(message)

    debug = info = warning = _log


class FakeStub(object):
    """Answers run_jsx_file like the AYON extension: the script's value wrapped as a result."""

    def __init__(self, spans):
        self.spans = spans  # comp id -> JSON string the JSX would return
        self.scripts = []

    def run_jsx_file(self, path):
        with open(path) as handle:
            script = handle.read()
        self.scripts.append(script)
        comp_id = int(script.rsplit("(", 1)[1].split(")")[0])
        return self.spans[comp_id]


def render_instance(comp_id, target="frames", **extra):
    data = {"productType": "render", "productName": "renderMain%d" % comp_id,
            "members": [comp_id], "creator_attributes": {"render_target": target}}
    data.update(extra)
    return FakeInstance(**data)


class FrameRangePluginTests(unittest.TestCase):
    def test_orders_sit_around_ayon_collectors(self):
        # CollectAERender is CollectorOrder + 0.125, CollectExistingFrames + 0.150
        self.assertLess(plugin.READ_ORDER, 0.125)
        self.assertTrue(0.125 < plugin.APPLY_ORDER < 0.150)

    def test_parse_time_span(self):
        self.assertEqual(plugin.parse_time_span('{"frameStart": 1001, "frameEnd": 1050}'), (1001, 1050))
        with self.assertRaises(ValueError) as caught:
            plugin.parse_time_span('{"error": "Expected 1 render queue item", "count": 2}')
        self.assertIn("Expected 1", str(caught.exception))
        with self.assertRaises(ValueError):
            plugin.parse_time_span('{"frameStart": 1001, "frameEnd": 1000}')

    def test_read_time_span_runs_jsx_for_the_comp_and_cleans_up(self):
        stub = FakeStub({42: json.dumps({"frameStart": 1010, "frameEnd": 1020})})
        self.assertEqual(plugin.read_time_span(stub, 42), (1010, 1020))
        self.assertIn("})(42);", stub.scripts[0])
        self.assertIn("timeSpanStart", stub.scripts[0])

    def test_frames_instances_get_the_render_queue_range(self):
        stub = FakeStub({1: json.dumps({"frameStart": 1010, "frameEnd": 1040}),
                         3: json.dumps({"error": "Expected 1 render queue item for 'C', found 2"})})
        frames = render_instance(1)
        local = render_instance(2, target="local")
        broken = render_instance(3)
        inactive = render_instance(4, active=False)
        other = FakeInstance(productType="workfile", members=[5])
        context = [frames, local, broken, inactive, other]

        ayon_module = types.ModuleType("ayon_aftereffects.api")
        ayon_module.get_stub = lambda: stub
        sys.modules["ayon_aftereffects"] = types.ModuleType("ayon_aftereffects")
        sys.modules["ayon_aftereffects.api"] = ayon_module
        try:
            reader = plugin.CollectTSRenderQueueTimeSpan()
            reader.log = FakeLog()
            reader.process(context)
        finally:
            del sys.modules["ayon_aftereffects"], sys.modules["ayon_aftereffects.api"]

        self.assertEqual(frames.data[plugin.RANGE_KEY], [1010, 1040])
        for instance in (local, broken, inactive, other):
            self.assertNotIn(plugin.RANGE_KEY, instance.data)
        self.assertTrue(any("found 2" in message for message in reader.log.messages))

        # CollectAERender then sets the comp range with handles
        frames.data.update(frameStart=1001, frameEnd=1100, handleStart=8, handleEnd=8,
                           frameStartHandle=993, frameEndHandle=1108)
        applier = plugin.ApplyTSRenderQueueFrameRange()
        applier.log = FakeLog()
        applier.process(context)
        self.assertEqual(
            {key: frames.data[key] for key in plugin.frame_range_data(0, 0)},
            {"frameStart": 1010, "frameEnd": 1040, "handleStart": 0, "handleEnd": 0,
             "frameStartHandle": 1010, "frameEndHandle": 1040})

    def test_render_queue_errors_keep_the_collected_range(self):
        instance = render_instance(1, frameStart=1001, frameEnd=1100, render_queue_error="duplicate")
        instance.data[plugin.RANGE_KEY] = [1010, 1040]
        applier = plugin.ApplyTSRenderQueueFrameRange()
        applier.log = FakeLog()
        applier.process([instance])
        self.assertEqual(instance.data["frameStart"], 1001)


class FakeCreated(dict):
    def __init__(self, creator_identifier, members):
        super(FakeCreated, self).__init__(members=members)
        self.creator_identifier = creator_identifier


class PlanTests(unittest.TestCase):
    def test_plan_splits_existing_and_missing(self):
        render_10 = FakeCreated("render", [10])
        render_10_again = FakeCreated("render", ["10"])
        workfile = FakeCreated("workfile", [11])
        render_99 = FakeCreated("render", [99])
        broken = FakeCreated("render", [])

        existing, missing = ayon_publish.plan(
            [render_10, render_10_again, workfile, render_99, broken], [10, 11, "12", 11])
        self.assertEqual(existing, {10: [render_10, render_10_again]})
        self.assertEqual(missing, [11, 12])


if __name__ == "__main__":
    unittest.main()
