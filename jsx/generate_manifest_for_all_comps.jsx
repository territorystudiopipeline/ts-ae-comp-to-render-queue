// Error logging helper
function logError(msg) {
    try {
        var logFile;
        if (app.project && app.project.file && app.project.file.parent) {
            logFile = new File(app.project.file.parent.fsName + "/ae_manifest_error.log");
        } else {
            logFile = new File(Folder.myDocuments.fsName + "/ae_manifest_error.log");
        }
        if (logFile.open("a")) {
            logFile.writeln((new Date()).toISOString() + " - " + msg);
            logFile.close();
        }
    } catch (e) {
        // If logging fails, there's not much we can do, but we don't want to throw an error from the logger itself
    }
}

// Safe file writing helper
function safeWriteFile(filePath, data) {
    var file = new File(filePath);
    if (!file.open("w")) {
        var errMsg = "Cannot open manifest file for writing: " + filePath;
        logError(errMsg);
        throw new Error(errMsg);
    }
    try {
        file.write(data);
    } catch (e) {
        logError("Error writing manifest: " + e.toString());
        throw e;
    } finally {
        file.close();
    }
}
/*
    generate_manifest_for_all_comps.jsx
    This script generates a manifest JSON file for all comps in the current After Effects project, including nested comps.
    The manifest includes details about each comp such as its name, frame rate, duration, fonts used, and effects applied
    (both native and third-party).

    The output manifest is saved as "project_manifest.json" in the same folder as the project file or in a specified
    output location if defined in a "_comp_identifiers.json" file.

    The script also supports a debug mode that can be enabled by setting the environment variable 'JSX_DEBUG' to "1".
    In debug mode, additional alerts will provide information about the script's execution and any issues encountered.
*/

// DEBUG flag, set by environment variable 'JSX_DEBUG' if present
var DEBUG = 0;
try {
    var envDebug = $.getenv ? $.getenv("JSX_DEBUG") : null;
    if (envDebug === "1") {
        DEBUG = 1;
    }
} catch (e) {
    DEBUG = 0;
}

function getCompOutputLocationsFromJson() {
    var compOutputMap = {};
    var projectFile = app.project.file;
    if (!projectFile) {
        var msg = "Project file is not saved. Please save your project before running this script.";
        logError(msg);
        if (DEBUG) alert(msg);
        return compOutputMap;
    }
    var projectFolder = projectFile.parent;
    var parentFolder = projectFolder.parent;
    var jsonFile = null;
    try {
        var allFiles = projectFolder.getFiles();
        for (var j = 0; j < allFiles.length; j++) {
            var fileObj1 = allFiles[j];
            if (fileObj1 instanceof File && fileObj1.name.match(/_comp_identifiers\.json$/i)) {
                jsonFile = fileObj1;
                break;
            }
        }
        if (!jsonFile) {
            var parentFiles = parentFolder.getFiles();
            for (var k = 0; k < parentFiles.length; k++) {
                var fileObj2 = parentFiles[k];
                if (fileObj2 instanceof File && fileObj2.name.match(/_comp_identifiers\.json$/i)) {
                    jsonFile = fileObj2;
                    break;
                }
            }
        }
        if (!jsonFile) {
            var msg2 = "_comp_identifiers.json not found in either project or parent folder.";
            logError(msg2);
            if (DEBUG) alert(msg2);
            return compOutputMap;
        }
        if (jsonFile.open("r")) {
            try {
                var jsonStr = jsonFile.read();
                var compsList = JSON.parse(jsonStr);
                if (compsList && compsList.length) {
                    for (var i = 0; i < compsList.length; i++) {
                        var entry = compsList[i];
                        if (entry.name && entry.output_location) {
                            compOutputMap[entry.name] = entry.output_location;
                        }
                    }
                    }
            } catch (e) {
                logError("Failed to parse comp identifiers JSON: " + e.toString());
                if (DEBUG) alert("Failed to parse comp identifiers JSON: " + e);
            }
            jsonFile.close();
        }
    } catch (e) {
        logError("Error reading comp identifiers JSON: " + e.toString());
    }
    return compOutputMap;
}

function findAllComps() {
    var comps = [];
    var items = app.project.items;
    for (var i = 1; i <= items.length; i++) {
        var item = items[i];
        if (item instanceof CompItem) {
            comps.push(item);
        }
    }
    return comps;
}

function containsFont(arr, fontObj) {
    for (var i = 0; i < arr.length; i++) {
        var f = arr[i];
        if (f.name === fontObj.name && f.family === fontObj.family && f.style === fontObj.style) {
            return true;
        }
    }
    return false;
}
function containsEffect(arr, effectObj) {
    for (var i = 0; i < arr.length; i++) {
        var e = arr[i];
        if (e.name === effectObj.name && e.matchName === effectObj.matchName) {
            return true;
        }
    }
    return false;
}
function isNativeEffect(matchName) {
    return typeof matchName === "string" && (matchName.indexOf("ADBE") === 0 || matchName.indexOf("CC") === 0);
}
function collectManifestData(comp, manifestCache, sceneFilePath) {
    if (!comp) {
        return {
            scene_file: sceneFilePath || null,
            comp_name: null,
            comp_id: null,
            comp_frame_rate: null,
            comp_duration: null,
            fonts: [],
            effects: {
                native_effects: [],
                third_party_effects: []
            },
            nested_comps: [],
        };
    }
    var compIdentifier = comp.name;
    if (manifestCache[compIdentifier]) {
        return manifestCache[compIdentifier];
    }
    var manifest = {
        scene_file: sceneFilePath || null,
        comp_name: comp.name,
        comp_id: null,
        comp_frame_rate: comp.frameRate,
        comp_duration: comp.duration,
        fonts: [],
        effects: {
            native_effects: [],
            third_party_effects: []
        },
        nested_comps: []
    };
    for (var i = 1; i <= comp.numLayers; i++) {
        var layer = comp.layer(i);
        if (layer.source && layer.source instanceof CompItem) {
            var nestedComp = layer.source;
            manifest.nested_comps.push({
                comp_name: nestedComp.name,
                comp_id: null
            });
            if (nestedComp) {
                var nestedManifest = collectManifestData(nestedComp, manifestCache, sceneFilePath);
                for (var f = 0; f < nestedManifest.fonts.length; f++) {
                    if (!containsFont(manifest.fonts, nestedManifest.fonts[f])) {
                        manifest.fonts.push(nestedManifest.fonts[f]);
                    }
                }
                for (var n = 0; n < nestedManifest.effects.native_effects.length; n++) {
                    if (!containsEffect(manifest.effects.native_effects, nestedManifest.effects.native_effects[n])) {
                        manifest.effects.native_effects.push(nestedManifest.effects.native_effects[n]);
                    }
                }
                for (var t = 0; t < nestedManifest.effects.third_party_effects.length; t++) {
                    if (!containsEffect(manifest.effects.third_party_effects, nestedManifest.effects.third_party_effects[t])) {
                        manifest.effects.third_party_effects.push(nestedManifest.effects.third_party_effects[t]);
                    }
                }
            }
        }
        if (layer instanceof TextLayer) {
            var textProp = layer.property("Source Text");
            if (textProp) {
                var textDocument = textProp.value;
                if (textDocument && textDocument.font) {
                    var fontInfo = {
                        name: textDocument.font,
                        family: textDocument.fontFamily || null,
                        style: textDocument.fontStyle || null
                    };
                    manifest.fonts.push(fontInfo);
                }
            }
        }
        var effectParade = layer.property("ADBE Effect Parade");
        if (effectParade) {
            for (var e = 1; e <= effectParade.numProperties; e++) {
                var effect = effectParade.property(e);
                var effectObj = {
                    name: effect.name,
                    matchName: effect.matchName
                };
                if (isNativeEffect(effect.matchName)) {
                    if (!containsEffect(manifest.effects.native_effects, effectObj)) {
                        manifest.effects.native_effects.push(effectObj);
                    }
                } else {
                    if (!containsEffect(manifest.effects.third_party_effects, effectObj)) {
                        manifest.effects.third_party_effects.push(effectObj);
                    }
                }
            }
        }
    }
    manifestCache[compIdentifier] = manifest;
    return manifest;
}
function writeProjectManifestFile(sceneFilePath, compManifests, outputFolderPath) {
    var manifest = {
        scene_file: sceneFilePath,
        comps: compManifests
    };
    var filePath = outputFolderPath + "/project_manifest.json";
    try {
        safeWriteFile(filePath, JSON.stringify(manifest, null, 4));
    } catch (e) {
        logError("Failed to write project manifest: " + e.toString());
        if (DEBUG) alert("Failed to write project manifest: " + e);
    }
}

function getProjectManifestOutputFolder() {
    var projectFile = app.project.file;
    if (!projectFile) return null;
    var projectFolder = projectFile.parent;
    var parentFolder = projectFolder.parent;
    var jsonFile = null;
    var allFiles = projectFolder.getFiles();
    for (var j = 0; j < allFiles.length; j++) {
        var fileObj3 = allFiles[j];
        if (fileObj3 instanceof File && fileObj3.name.match(/_comp_identifiers\.json$/i)) {
            jsonFile = fileObj3;
            break;
        }
    }
    if (!jsonFile) {
        var parentFiles = parentFolder.getFiles();
        for (var k = 0; k < parentFiles.length; k++) {
            var fileObj4 = parentFiles[k];
            if (fileObj4 instanceof File && fileObj4.name.match(/_comp_identifiers\.json$/i)) {
                jsonFile = fileObj4;
                break;
            }
        }
    }
    if (!jsonFile) return projectFolder.fsName;
    if (jsonFile.open("r")) {
        try {
            var jsonStr = jsonFile.read();
            var compsList = JSON.parse(jsonStr);
            if (compsList.length && compsList[0].output_location) {
                return compsList[0].output_location;
            }
        } catch (e) {
            logError("Failed to parse comp identifiers JSON for manifest output location: " + e.toString());
            if (DEBUG) alert("Failed to parse comp identifiers JSON for manifest output location: " + e);
        }
        jsonFile.close();
    }
    return projectFolder.fsName;
}

function main() {
    if (!app.project || !app.project.file) {
        var msg = "No project is open. Please open a project before running this script.";
        logError(msg);
        if (DEBUG) alert(msg);
        return;
    }
    var compOutputMap = getCompOutputLocationsFromJson();
    var allComps = findAllComps();
    if (!allComps.length) {
        var msg2 = "No comps found in the project.";
        logError(msg2);
        if (DEBUG) alert(msg2);
        return;
    }
    var manifestCache = {};
    var sceneFileName = app.project.file.name;
    var projectFolder = app.project.file.parent.fsName;
    var sceneFilePath = (new File(projectFolder + "/" + sceneFileName)).fsName;
    var compManifests = {};
    for (var i = 0; i < allComps.length; i++) {
        var comp = allComps[i];
        if (!comp) {
            logError("Null comp encountered in allComps at index " + i);
            continue;
        }
        var outputFolderPath = compOutputMap[comp.name];
        if (!outputFolderPath) {
            outputFolderPath = projectFolder;
        }
        var compSceneFilePath = (new File(outputFolderPath + "/" + sceneFileName)).fsName;
        var manifest = collectManifestData(comp, manifestCache, compSceneFilePath);
        delete manifest.scene_file;
        compManifests[comp.name] = manifest;
    }
    var manifestOutputFolder = getProjectManifestOutputFolder();
    writeProjectManifestFile(sceneFilePath, compManifests, manifestOutputFolder);
    if (DEBUG) alert("Project manifest file created at: " + manifestOutputFolder);
}

try {
    main();
} catch (e) {
    logError("Manifest generation failed: " + e.toString() + (e.line ? " (Line: " + e.line + ")" : ""));
    if (DEBUG) alert("Manifest generation failed. See ae_manifest_error.log in your Documents folder.");
    throw e;
}
