// SPIKE step 1: what can scripting see and set in this After Effects (colour / proxies / settings)
#include "common.jsxinc"
(function () {
    var A = $.global.TS_JSX_ARGS;
    var R = {steps: [], version: app.version, buildName: app.buildName};
    var proj = app.project;
    R.project = spStep(R, "project reflect", function () { return spReflect(proj); });
    R.projectValues = spStep(R, "project colour values", function () { return spValues(proj, SP_COLOUR); });

    app.beginUndoGroup("ts proxy spike inspect");
    var comp = null, rq = null;
    try {
        comp = proj.items.addComp("__ts_spike_inspect", 64, 64, 1, 1, 25);
        var solidLayer = comp.layers.addSolid([0.5, 0.5, 0.5], "__ts_spike_solid", 64, 64, 1);
        var footage = solidLayer.source;
        R.compItem = spStep(R, "CompItem reflect", function () { return spReflect(comp); });
        R.footageItem = spStep(R, "FootageItem reflect", function () { return spReflect(footage); });
        R.footageSource = spStep(R, "FootageSource reflect (solid)", function () { return spReflect(footage.mainSource); });
        R.footageSourceValues = spStep(R, "FootageSource values", function () { return spValues(footage.mainSource, SP_COLOUR); });
        R.footageItemValues = spStep(R, "FootageItem values", function () { return spValues(footage, SP_COLOUR); });

        rq = proj.renderQueue.items.add(comp);
        var om = rq.outputModule(1);
        R.rqItem = spStep(R, "RenderQueueItem reflect", function () { return spReflect(rq); });
        R.outputModule = spStep(R, "OutputModule reflect", function () { return spReflect(om); });
        R.rqTemplates = spStep(R, "render settings templates", function () { return rq.templates; });
        R.omTemplates = spStep(R, "output module templates", function () { return om.templates; });
        R.rqSettable = spStep(R, "render settings (settable)", function () { return rq.getSettings(GetSettingsFormat.STRING_SETTABLE); });
        R.omSettable = spStep(R, "output module settings (settable)", function () { return om.getSettings(GetSettingsFormat.STRING_SETTABLE); });
        R.omSettings = spStep(R, "output module settings (all)", function () { return om.getSettings(GetSettingsFormat.STRING); });
        spStep(R, "set Proxy Use = Use Comp Proxies Only", function () {
            rq.setSetting("Proxy Use", "Use Comp Proxies Only");
            return rq.getSettings(GetSettingsFormat.STRING)["Proxy Use"];
        });
    } finally {
        try { if (rq) rq.remove(); } catch (e) {}
        try {
            for (var i = proj.numItems; i >= 1; i--) {
                var it = proj.item(i);
                if (it.name.indexOf("__ts_spike") === 0) it.remove();
            }
        } catch (e2) {}
        app.endUndoGroup();
    }
    spWrite(A.out, R);
})();
