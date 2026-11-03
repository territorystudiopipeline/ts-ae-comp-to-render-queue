// SPIKE step 2: render a precomp to EXR and PNG, use each as its comp proxy, render hosts that use it.
// Run in a NEW, EMPTY project. Saves it to <folder>/proxy_spike.aep for step 3.
#include "common.jsxinc"
(function () {
    var A = $.global.TS_JSX_ARGS;
    var R = {steps: [], version: app.version};
    var proj = app.project;
    if (proj.numItems > 0) throw new Error("Run step 2 in a new, empty project (File > New > New Project).");
    var folder = A.folder, fps = A.fps, n = A.frames, s = A.size, testFrame = A.testFrame;
    new Folder(folder + "/proxies/exr").create();
    new Folder(folder + "/proxies/png").create();
    new Folder(folder + "/verify").create();
    new Folder(folder + "/farm").create();

    spStep(R, "project 32 bpc", function () { proj.bitsPerChannel = 32; return proj.bitsPerChannel; });
    R.projectValues = spValues(proj, SP_COLOUR);

    // Precomp: a box moving 8px per frame (frame alignment), two solids added together so values
    // go above 1.0 (0.36, 1.0, 1.8) - tests EXR float and PNG clipping. Transparent around it (alpha).
    var P = proj.items.addComp("P_precomp", s, s, 1, n / fps, fps);
    var solids = [];
    for (var k = 0; k < 2; k++) {
        var L = P.layers.addSolid([0.18, 0.5, 0.9], "box" + k, s / 2, s / 2, 1);
        L.property("ADBE Transform Group").property("ADBE Position").expression =
            "[" + (s / 4) + " + timeToFrames(time) * 8, " + (s / 2) + "]";
        if (k === 1) L.blendingMode = BlendingMode.ADD;
        solids.push(L);
    }
    R.precomp = {name: P.name, frameRate: P.frameRate, duration: P.duration, frames: n};

    function host(name, setup) {
        var h = proj.items.addComp(name, s, s, 1, n / fps, fps);
        var layer = h.layers.add(P);
        if (setup) setup(layer);
        return h;
    }
    var hosts = {
        H_normal: host("H_normal"),
        H_collapsed: host("H_collapsed", function (l) { l.collapseTransformation = true; }),
        H_3d: host("H_3d", function (l) { l.threeDLayer = true; })
    };

    function applyFormat(om, fmt) {
        var info = {fmt: fmt};
        var pattern = fmt === "exr" ? /exr/i : /png/i;
        var templates = om.templates;
        for (var i = 0; i < templates.length; i++) {
            if (pattern.test(templates[i])) { om.applyTemplate(templates[i]); info.template = templates[i]; break; }
        }
        if (!info.template) {
            try { om.setSettings({"Format": fmt === "exr" ? "OpenEXR Sequence" : "PNG Sequence"}); info.setFormat = "ok"; }
            catch (e) { info.setFormat = String(e); }
        }
        try { om.setSettings({"Channels": "RGB + Alpha"}); info.channels = "ok"; } catch (e2) { info.channels = String(e2); }
        if (fmt === "png") {
            try { om.setSettings({"Depth": "Trillions of Colors+"}); info.depth = "ok"; }
            catch (e3) { try { om.setSettings({"Depth": "Trillions of Colors"}); info.depth = "ok (no +)"; } catch (e4) { info.depth = String(e4); } }
        }
        try { info.settings = om.getSettings(GetSettingsFormat.STRING); } catch (e5) { info.settingsError = String(e5); }
        return info;
    }
    function queue(comp, file, fmt, start, count, proxyUse) {
        var rq = proj.renderQueue.items.add(comp);
        rq.timeSpanStart = start / comp.frameRate;
        rq.timeSpanDuration = count / comp.frameRate;
        var info = {comp: comp.name, file: file};
        if (proxyUse) {
            try { rq.setSetting("Proxy Use", proxyUse); info.proxyUse = rq.getSettings(GetSettingsFormat.STRING)["Proxy Use"]; }
            catch (e) { info.proxyUseError = String(e); }
        } else {
            try { info.proxyUse = rq.getSettings(GetSettingsFormat.STRING)["Proxy Use"]; } catch (e1) {}
        }
        var om = rq.outputModule(1);
        info.format = applyFormat(om, fmt);
        om.file = new File(file);
        return info;
    }
    function clearQueue() {
        for (var i = proj.renderQueue.numItems; i >= 1; i--) proj.renderQueue.item(i).remove();
    }
    function render(label) {
        var t = new Date().getTime();
        proj.renderQueue.render();
        var statuses = [];
        for (var i = 1; i <= proj.renderQueue.numItems; i++) statuses.push(String(proj.renderQueue.item(i).status));
        R.steps.push({name: "render " + label, ok: true, value: {seconds: (new Date().getTime() - t) / 1000, statuses: statuses}});
    }

    // 1. Proxies: the whole precomp, EXR and PNG
    R.proxyRenders = [
        spStep(R, "queue exr proxy", function () { return queue(P, folder + "/proxies/exr/P_precomp_[####].exr", "exr", 0, n); }),
        spStep(R, "queue png proxy", function () { return queue(P, folder + "/proxies/png/P_precomp_[####].png", "png", 0, n); })
    ];
    render("proxies");
    clearQueue();

    // 2. Change the precomp (box turns red) so renders show whether the proxy or the live comp was used
    for (var j = 0; j < solids.length; j++) solids[j].source.mainSource.color = [1, 0, 0];

    // 3. Each proxy format: set it, interpret it, render every host one frame
    function firstFrame(dir, ext) {
        var files = new Folder(dir).getFiles("*." + ext);
        files.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
        return files.length ? files[0] : null;
    }
    R.proxies = {};
    var formats = ["exr", "png"];
    for (var f = 0; f < formats.length; f++) {
        var fmt = formats[f];
        var info = {};
        R.proxies[fmt] = info;
        var first = firstFrame(folder + "/proxies/" + fmt, fmt);
        info.firstFrame = first ? first.fsName : null;
        if (!first) { info.error = "no proxy frames rendered"; continue; }
        spStep(R, fmt + ": setProxyWithSequence", function () { P.setProxyWithSequence(first, false); return P.hasProxy; });
        var ps = P.proxySource;
        info.reflect = spReflect(ps);
        info.before = spValues(ps, SP_COLOUR);
        spStep(R, fmt + ": conformFrameRate", function () { ps.conformFrameRate = fps; return ps.conformFrameRate; });
        spStep(R, fmt + ": alphaMode premultiplied", function () { ps.alphaMode = AlphaMode.PREMULTIPLIED; return String(ps.alphaMode); });
        // Optional colour settings to try, from step 1's findings: {"propertyName": value}
        var tries = (A.colourTries || {})[fmt] || {};
        for (var key in tries) {
            (function (k2, v2) {
                spStep(R, fmt + ": set proxy " + k2, function () { ps[k2] = v2; return ps[k2]; });
            })(key, tries[key]);
        }
        spStep(R, fmt + ": useProxy", function () { P.useProxy = true; return P.useProxy; });
        info.after = spValues(ps, SP_COLOUR);
        info.duration = ps.duration !== undefined ? ps.duration : null;
        info.renders = [];
        for (var hn in hosts) {
            info.renders.push(queue(hosts[hn], folder + "/verify/" + fmt + "_" + hn + "_[####].exr", "exr", testFrame, 1,
                                    "Use Comp Proxies Only"));
        }
        if (fmt === "exr") {
            // Default render settings: does a plain render (as aerender -comp would add) use the proxy?
            info.renders.push(queue(hosts.H_normal, folder + "/verify/exr_default_settings_H_normal_[####].exr", "exr", testFrame, 1, null));
        }
        render(fmt + " hosts");
        clearQueue();
    }

    // 4. Reference: the live (now red) precomp, no proxy
    P.useProxy = false;
    R.live = queue(hosts.H_normal, folder + "/verify/live_H_normal_[####].exr", "exr", testFrame, 1, "Use No Proxies");
    render("live reference");
    clearQueue();

    // 5. Scene for the farm test: no proxy set, one queued item (H_normal, test frame); the patch sets the rest
    P.setProxyToNone();
    R.farm = queue(hosts.H_normal, folder + "/farm/farm_H_normal_[####].exr", "exr", testFrame, 1, null);
    spStep(R, "save proxy_spike.aep", function () { proj.save(new File(folder + "/proxy_spike.aep")); return proj.file.fsName; });
    spWrite(A.out, R);
})();
