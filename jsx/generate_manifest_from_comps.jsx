/*
    generate_manifest_from_comps.jsx
    This script generates a manifest JSON file for specified comps in an After Effects project. The manifest includes:
- comp name
- comp frame rate
- comp duration
- fonts used (name, family, style)
- effects used (separated into native and third-party based on matchName)
- nested comps (name only, since id is not accessible)

The script looks for a JSON file named "_comp_identifiers.json" in the project folder or its parent folder to get the list of comps to analyze.
The JSON should be an array of objects with the following structure:
[
    {
        "name": "Comp Name",
        "id": "Optional ID (not used for matching due to AE limitations)",
        "output_location": "Optional folder path to save the manifest (defaults to project folder)"
    },
    ...
]

For each comp specified, the script generates a manifest JSON file named "<comp_name>_manifest.json" in the specified output location or the project folder if not specified.


ALl debugging alerts have been commented out for cleaner execution, but can be uncommented for troubleshooting if needed.
*/

// DEBUG flag, set by environment variable 'JSX_DEBUG' if present
var DEBUG = 0;
try {
    var envDebug = $.getenv ? $.getenv("JSX_DEBUG") : null;
    if (envDebug === "1") {
        DEBUG = 1;
    }
} catch (e) {
    // If getenv is not available, default to 0
    DEBUG = 0;
}

// Try to read comp identifiers from a JSON file in the project folder
function getCompsToAnalyzeFromJson() {
    var compsList = null;
    var projectFile = app.project.file;
    if (!projectFile) {
        if (DEBUG) alert("Project file is not saved. Please save your project before running this script.");
        return null;
    }
    var projectFolder = projectFile.parent;
    var parentFolder = projectFolder.parent;
    var jsonFile = null;
    // List all files in projectFolder for debugging
    var allFiles = projectFolder.getFiles();
    var allNames = [];
    for (var j = 0; j < allFiles.length; j++) {
        allNames.push(allFiles[j].fsName);
     }
     if (DEBUG) alert("Files in project folder (" + projectFolder.fsName + "):\n" + allNames.join("\n"));
     if (DEBUG) alert("Files in parent folder (" + parentFolder.fsName + "):\n" + parentNames.join("\n"));

     // Search for _comp_identifiers.json manually
    for (var j = 0; j < allFiles.length; j++) {
        var fileObj = allFiles[j];
        if (fileObj instanceof File && fileObj.name.match(/_comp_identifiers\.json$/i)) {
            jsonFile = fileObj;
            break;
        }
    }
    // If not found, try parent folder
    if (!jsonFile) {
        var parentFiles = parentFolder.getFiles();
        var parentNames = [];
         for (var k = 0; k < parentFiles.length; k++) {
             parentNames.push(parentFiles[k].fsName);
         }
         if (DEBUG) alert("Files in parent folder (" + parentFolder.fsName + "):\n" + parentNames.join("\n"));
        for (var k = 0; k < parentFiles.length; k++) {
            var fileObj = parentFiles[k];
            if (fileObj instanceof File && fileObj.name.match(/_comp_identifiers\.json$/i)) {
                jsonFile = fileObj;
                break;
            }
        }
    }
    if (!jsonFile) {
        if (DEBUG) alert("_comp_identifiers.json not found in either project or parent folder.");
        return null;
    }  else {
        if (DEBUG) alert("Found _comp_identifiers.json at: " + jsonFile.fsName);
     }
    if (jsonFile.open("r")) {
        try {
            var jsonStr = jsonFile.read();
            compsList = JSON.parse(jsonStr);
        } catch (e) {
            if (DEBUG) alert("Failed to parse comp identifiers JSON: " + e);
        }
        jsonFile.close();
    }
    return compsList;
}

var compsToAnalyze = getCompsToAnalyzeFromJson();



function findCompByNameAndId(name, id) {
    var items = app.project.items;
    for (var i = 1; i <= items.length; i++) {
        var item = items[i];
        if (item instanceof CompItem && item.name === name) {
            // AE doesn't expose comp.id by default, so match by name only
            return item;
        }
    }
    return null;
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
    // Consider effects native if matchName starts with "ADBE" or "CC"
    return typeof matchName === "string" && (matchName.indexOf("ADBE") === 0 || matchName.indexOf("CC") === 0);

}

function collectManifestData(comp, manifestCache, sceneFilePath) {
    if (!comp) {
        // Defensive: should never happen, but return empty manifest if so
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
        comp_id: null, // AE doesn't expose comp.id
        comp_frame_rate: comp.frameRate,
        comp_duration: comp.duration,
        fonts: [],
        effects: {
            native_effects: [],
            third_party_effects: []
        },
        nested_comps: [],
    };
    for (var i = 1; i <= comp.numLayers; i++) {
        var layer = comp.layer(i);
        // Nested comps
        if (layer.source && layer.source instanceof CompItem) {
            var nestedComp = layer.source;
            manifest.nested_comps.push({
                comp_name: nestedComp.name,
                comp_id: null
            });
            if (nestedComp) {
                var nestedManifest = collectManifestData(nestedComp, manifestCache);
                // Merge fonts/plugins from nested
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
        // Text layers
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
        // Effects
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

function writeManifestFile(manifest, folderPath) {
    var fileName = manifest.comp_name + "_manifest.json";
    var folderString = (folderPath instanceof Folder) ? folderPath.fsName : folderPath;
    var file = new File(folderString + "/" + fileName);
    file.open("w");
    file.write(JSON.stringify(manifest, null, 4));
    file.close();
}

function main() {
    if (!app.project || !app.project.file) {
        if (DEBUG) alert("No project is open. Please open a project before running this script.");
        return;
    }
    if (!compsToAnalyze || !compsToAnalyze.length) {
        if (DEBUG) alert("No comps to analyze. Please check your _comp_identifiers.json file or input list.");
        return;
    }
    var manifestCache = {};
    var sceneFileName = app.project.file.name;
    for (var i = 0; i < compsToAnalyze.length; i++) {
        var compInfo = compsToAnalyze[i];
        var comp = findCompByNameAndId(compInfo.name, compInfo.id);
        var outputFolderPath = compInfo.output_location;
        if (!outputFolderPath) {
            if (DEBUG) alert("No output_location specified for comp '" + compInfo.name + "'. Manifest will be saved to project folder.");
            outputFolderPath = app.project.file.parent.fsName;
        }
        // Use File/Folder to join paths for OS-agnostic separator
        var sceneFilePath = (new File(outputFolderPath + "/" + sceneFileName)).fsName;
        if (comp) {
            var manifest = collectManifestData(comp, manifestCache, sceneFilePath);
            writeManifestFile(manifest, outputFolderPath);
        } // else {
            // alert("Comp not found: '" + compInfo.name + "' (id: " + compInfo.id + ")\nCheck that the comp exists in the project and the name matches exactly.\nAll comp names in the project: " + getAllCompNames().join(", "));
        // }
    }
    // alert("Manifest files created.");
}

function getAllCompNames() {
    var names = [];
    var items = app.project.items;
    for (var i = 1; i <= items.length; i++) {
        if (items[i] instanceof CompItem) {
            names.push(items[i].name);
        }
    }
    return names;
}

main();
