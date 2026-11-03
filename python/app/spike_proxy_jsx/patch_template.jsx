// SPIKE step 3 (farm, AfterFX -noui -r): open the scene copy, set the precomp's proxy, save a patched copy.
// Generated from patch_template.jsx with the paths filled in.
var SPIKE = %(config)s;
#include "%(common)s"
(function () {
    var R = {steps: [], version: app.version, host: $.getenv("COMPUTERNAME"), started: String(new Date())};
    try {
        app.beginSuppressDialogs();
        spStep(R, "open scene", function () { app.open(new File(SPIKE.scene)); return app.project.file.fsName; });
        var P = spCompByName(SPIKE.precomp);
        if (!P) throw new Error("precomp not found: " + SPIKE.precomp);
        spStep(R, "setProxyWithSequence", function () { P.setProxyWithSequence(new File(SPIKE.firstFrame), false); return P.hasProxy; });
        var ps = P.proxySource;
        spStep(R, "conformFrameRate", function () { ps.conformFrameRate = P.frameRate; return ps.conformFrameRate; });
        spStep(R, "alphaMode", function () { ps.alphaMode = AlphaMode.PREMULTIPLIED; return String(ps.alphaMode); });
        var tryKeys = spKeys(SPIKE.colourTries);
        for (var t = 0; t < tryKeys.length; t++) {
            (function (k, v) { spStep(R, "set proxy " + k, function () { ps[k] = v; return ps[k]; }); })(tryKeys[t], SPIKE.colourTries[tryKeys[t]]);
        }
        spStep(R, "useProxy", function () { P.useProxy = true; return P.useProxy; });
        R.proxyValues = spValues(ps, SP_COLOUR);
        spStep(R, "Proxy Use on the queued item", function () {
            var q = app.project.renderQueue, done = [];
            for (var i = 1; i <= q.numItems; i++) {
                var it = q.item(i);
                if (it.comp.name === SPIKE.host) { it.setSetting("Proxy Use", "Use Comp Proxies Only"); done.push(i); }
            }
            return done;
        });
        spStep(R, "save patched", function () { app.project.save(new File(SPIKE.patched)); return app.project.file.fsName; });
        R.ok = true;
    } catch (e) {
        R.ok = false;
        R.error = String(e) + (e.line ? " (line " + e.line + ")" : "");
    } finally {
        R.finished = String(new Date());
        try { spWrite(SPIKE.result, R); } catch (e2) {}
        try { app.endSuppressDialogs(false); } catch (e3) {}
        app.quit();
    }
})();
