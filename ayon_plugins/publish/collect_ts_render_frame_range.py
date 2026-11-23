"""
Publish the frame range the render queue item actually rendered (usually the work area),
not the whole comp.

AYON's CollectAERender sets frameStart/frameEnd from the comp's duration and ignores the render
queue item's Time Span. For instances that publish existing frames (render target "frames", the
ones our render queue tool creates), that makes CollectExistingFrames expect frames that were never
rendered. These two plug-ins fix that:

1. CollectTSRenderQueueTimeSpan (before CollectAERender) reads each render instance's render queue
   item Time Span and stores it on the instance. It runs first because CollectAERender asks After
   Effects for render info, which re-queues Done items.
2. ApplyTSRenderQueueFrameRange (after CollectAERender, before CollectExistingFrames) replaces the
   comp range with that Time Span, with no handles.

The rez package adds this folder to PYBLISHPLUGINPATH.
"""
import json
import os
import tempfile

import pyblish.api

# Ordered around ayon_aftereffects' CollectAERender (+0.125) and CollectExistingFrames (+0.150).
READ_ORDER = pyblish.api.CollectorOrder + 0.12
APPLY_ORDER = pyblish.api.CollectorOrder + 0.13

RANGE_KEY = "tsRenderQueueFrameRange"
TARGETS = ("frames",)  # render targets whose frame range comes from the render queue item

# Reads one comp's render queue item Time Span as display frame numbers (the numbers After Effects
# writes into [####]). %d is the comp id. The value of the last expression is returned by
# $.evalFile, which the AYON extension wraps as {"result": ...}.
TIME_SPAN_JSX = """(function (compId) {
    var comp = app.project.itemByID(compId);
    if (!comp) {
        return JSON.stringify({"error": "There is no composition with id " + compId});
    }
    var queue = app.project.renderQueue;
    var items = [];
    for (var i = 1; i <= queue.numItems; ++i) {
        var item = queue.item(i);
        if (item.comp && item.comp.id == compId) {
            items.push(item);
        }
    }
    if (items.length != 1) {
        return JSON.stringify({"error": "Expected 1 render queue item for '" + comp.name +
                                        "', found " + items.length, "count": items.length});
    }
    var start = comp.displayStartFrame + Math.round(items[0].timeSpanStart / comp.frameDuration);
    var count = Math.round(items[0].timeSpanDuration / comp.frameDuration);
    return JSON.stringify({"frameStart": start, "frameEnd": start + count - 1});
})(%d);
"""


def is_render_instance(instance):
    data = instance.data
    if not data.get("active", True):
        return False
    return (data.get("productBaseType") or data.get("productType")) == "render"


def render_target(instance):
    return (instance.data.get("creator_attributes") or {}).get("render_target")


def parse_time_span(result):
    """
    ``(frameStart, frameEnd)`` from the JSX result, or raise ValueError with After Effects' reason.
    """
    data = json.loads(result) if isinstance(result, str) else result
    if not isinstance(data, dict):
        raise ValueError("Unexpected answer from After Effects: %r" % (result,))
    if data.get("error"):
        raise ValueError(data["error"])
    start, end = int(data["frameStart"]), int(data["frameEnd"])
    if end < start:
        raise ValueError("The render queue item renders no frames (%s-%s)" % (start, end))
    return start, end


def read_time_span(stub, comp_id):
    """Run TIME_SPAN_JSX for one comp through the AYON After Effects stub."""
    handle, path = tempfile.mkstemp(suffix=".jsx", prefix="ts_time_span_")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as script:
            script.write(TIME_SPAN_JSX % int(comp_id))
        return parse_time_span(stub.run_jsx_file(path))
    finally:
        os.remove(path)


def frame_range_data(frame_start, frame_end):
    """Instance data for a range rendered without handles."""
    return {
        "frameStart": frame_start,
        "frameEnd": frame_end,
        "handleStart": 0,
        "handleEnd": 0,
        "frameStartHandle": frame_start,
        "frameEndHandle": frame_end,
    }


class CollectTSRenderQueueTimeSpan(pyblish.api.ContextPlugin):
    """Read the render queue item Time Span of each existing-frames render instance."""

    order = READ_ORDER
    label = "Collect Render Queue Time Span (Territory)"
    hosts = ["aftereffects"]

    def process(self, context):
        from ayon_aftereffects.api import get_stub

        stub = None
        for instance in context:
            if not is_render_instance(instance) or render_target(instance) not in TARGETS:
                continue
            members = instance.data.get("members") or []
            if not members:
                continue
            stub = stub or get_stub()
            try:
                frame_range = read_time_span(stub, int(members[0]))
            except ValueError as error:
                # Missing or duplicate queue items are reported by AYON's ValidateRenderQueue;
                # the comp range stays in place.
                self.log.warning("%s: %s" % (instance.data.get("productName"), error))
                continue
            instance.data[RANGE_KEY] = list(frame_range)
            self.log.debug("%s renders frames %s-%s" % ((instance.data.get("productName"),) + frame_range))


class ApplyTSRenderQueueFrameRange(pyblish.api.ContextPlugin):
    """Replace the comp frame range collected by CollectAERender with the Time Span read earlier."""

    order = APPLY_ORDER
    label = "Apply Render Queue Frame Range (Territory)"
    hosts = ["aftereffects"]

    def process(self, context):
        for instance in context:
            frame_range = instance.data.get(RANGE_KEY)
            if not frame_range or instance.data.get("render_queue_error"):
                continue
            before = (instance.data.get("frameStartHandle"), instance.data.get("frameEndHandle"))
            instance.data.update(frame_range_data(*frame_range))
            self.log.info("%s: frame range %s-%s (comp range was %s-%s)"
                          % ((instance.data.get("productName"),) + tuple(frame_range) + before))
