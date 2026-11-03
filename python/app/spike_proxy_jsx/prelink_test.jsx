// SPIKE step 4: proxies linked BEFORE the save, to placeholder frames the farm overwrites.
// Run in a NEW, EMPTY project. Saves <folder>/prelink/prelink.aep for the farm jobs.
//
// Two precomps, P_off and P_on (red box, live). Each has its proxy set to placeholder frames
// (black) that a farm job overwrites by rendering SRC_off / SRC_on (green box) to the same
// path. P_off's proxy switch is off, P_on's is on. Hosts H_off / H_on render with
// "Use Comp Proxies Only". Farm output per frame: green = proxy used, red = live comp,
// black = stale placeholder. The box moves 8 px per frame (alignment).
#include "common.jsxinc"
(function () {
    var A = $.global.TS_JSX_ARGS;
    var R = {steps: [], version: app.version};
    var proj = app.project;
    if (proj.numItems > 0) throw new Error("Run step 4 in a new, empty project (File > New > New Project).");
    var root = A.folder + "/prelink", fps = A.fps, n = A.frames, s = A.size;
    var dirs = ["/placeholder", "/proxies/P_off", "/proxies/P_on", "/out"];
    for (var d = 0; d < dirs.length; d++) new Folder(root + dirs[d]).create();
    proj.bitsPerChannel = 32;

    function boxComp(name, colour) {
        var c = proj.items.addComp(name, s, s, 1, n / fps, fps);
        var L = c.layers.addSolid(colour, "box", s / 2, s / 2, 1);
        L.property("ADBE Transform Group").property("ADBE Position").expression =
            "[" + (s / 4) + " + timeToFrames(time) * 8, " + (s / 2) + "]";
        return c;
    }
    function exrTemplate(om) {
        var t = om.templates;
        for (var i = 0; i < t.length; i++) if (/^exr$/i.test(t[i])) { om.applyTemplate(t[i]); return t[i]; }
        for (var j = 0; j < t.length; j++) if (/exr/i.test(t[j])) { om.applyTemplate(t[j]); return t[j]; }
        om.setSettings({"Format": "OpenEXR Sequence"});
        return "OpenEXR Sequence (set)";
    }
    function queue(comp, file, start, count, proxyUse) {
        var rq = proj.renderQueue.items.add(comp);
        rq.timeSpanStart = start * comp.frameDuration;   // exact frame times
        rq.timeSpanDuration = count * comp.frameDuration;
        rq.setSetting("Proxy Use", proxyUse);
        var om = rq.outputModule(1);
        var template = exrTemplate(om);
        om.file = new File(file);
        return {comp: comp.name, file: file, proxyUse: rq.getSettings(GetSettingsFormat.STRING)["Proxy Use"], template: template};
    }

    // 1. One black frame at the precomp's size, the placeholder for every proxy frame
    var blank = proj.items.addComp("__placeholder", s, s, 1, 1 / fps, fps);
    R.placeholderRender = spStep(R, "queue placeholder", function () { return queue(blank, root + "/placeholder/ph_[####].exr", 0, 1, "Use No Proxies"); });
    var t0 = new Date().getTime();
    proj.renderQueue.render();
    R.placeholderSeconds = (new Date().getTime() - t0) / 1000;
    proj.renderQueue.item(1).remove();
    var ph = new Folder(root + "/placeholder").getFiles("*.exr");
    if (!ph.length) throw new Error("placeholder frame wasn't rendered");
    blank.remove();

    // 2. Live precomps (red), proxy sources (green), hosts
    var setups = [{name: "P_off", switchOn: false}, {name: "P_on", switchOn: true}];
    R.setups = [];
    for (var k = 0; k < setups.length; k++) {
        var cfg = setups[k];
        var P = boxComp(cfg.name, [1, 0, 0]);
        var SRC = boxComp("SRC_" + cfg.name.split("_")[1], [0, 1, 0]);
        var H = proj.items.addComp("H_" + cfg.name.split("_")[1], s, s, 1, n / fps, fps);
        H.layers.add(P);
        var proxyDir = root + "/proxies/" + cfg.name;
        var info = {precomp: P.name, source: SRC.name, host: H.name, switchOn: cfg.switchOn};
        // Placeholder frames with the names the proxy render will write (frame 0..n-1)
        info.copies = spStep(R, cfg.name + ": placeholder frames", function () {
            var made = 0;
            for (var f = 0; f < n; f++) {
                var num = String(10000 + f).substr(1);
                if (ph[0].copy(proxyDir + "/" + cfg.name + "_" + num + ".exr")) made++;
            }
            return made;
        });
        spStep(R, cfg.name + ": setProxyWithSequence", function () {
            P.setProxyWithSequence(new File(proxyDir + "/" + cfg.name + "_0000.exr"), false);
            return P.hasProxy;
        });
        spStep(R, cfg.name + ": interpret", function () {
            P.proxySource.conformFrameRate = fps;
            P.proxySource.alphaMode = AlphaMode.PREMULTIPLIED;
            return {displayFrameRate: P.proxySource.displayFrameRate, alphaMode: String(P.proxySource.alphaMode)};
        });
        spStep(R, cfg.name + ": useProxy " + cfg.switchOn, function () { P.useProxy = cfg.switchOn; return P.useProxy; });
        // Proxy job item: the green source rendered over the placeholders
        info.proxyItem = spStep(R, cfg.name + ": queue proxy item", function () {
            return queue(SRC, proxyDir + "/" + cfg.name + "_[####].exr", 0, n, "Use No Proxies");
        });
        // Main job item: the host, proxies only
        info.mainItem = spStep(R, cfg.name + ": queue main item", function () {
            return queue(H, root + "/out/" + H.name + "_[####].exr", 0, n, "Use Comp Proxies Only");
        });
        R.setups.push(info);
    }
    spStep(R, "save prelink.aep", function () { proj.save(new File(root + "/prelink.aep")); return proj.file.fsName; });
    spWrite(A.out, R);
})();
