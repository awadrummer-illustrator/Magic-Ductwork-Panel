#target illustrator

// Define debug log early so we can use it
if (typeof MDUX === "undefined") {
    var MDUX = {};
}

function MDUX_debugLog_Early(message) {
    try {
        var f = new File($.fileName);
        var folder = f.parent.parent; // Extension root
        var logFile = new File(folder.fsName + "/debug.log");
        logFile.open("a");
        logFile.writeln("[EARLY] " + message);
        logFile.close();
    } catch(e) {}
}

MDUX_debugLog_Early("panel-bridge.jsx is loading...");

#include "./register-ignore.jsx"
#include "./export-utils.jsx"

MDUX_debugLog_Early("Includes finished. MDUX_performExport type: " + typeof MDUX_performExport);

// Store the jsx folder path when this file is first evaluated
// ALWAYS use $.global.MDUX_LAST_BRIDGE_PATH (set by panel.js) if available
// because $.fileName is unreliable and may point to system32 or other wrong locations
// Force update every time if MDUX_LAST_BRIDGE_PATH is set to override any stale cached values
if ($.global.MDUX_LAST_BRIDGE_PATH) {
    $.global.MDUX_JSX_FOLDER = File($.global.MDUX_LAST_BRIDGE_PATH).parent;
    try { $.writeln("[MDUX] Updated MDUX_JSX_FOLDER from MDUX_LAST_BRIDGE_PATH: " + $.global.MDUX_JSX_FOLDER); } catch(e) {}
} else if (typeof $.global.MDUX_JSX_FOLDER === "undefined") {
    $.global.MDUX_JSX_FOLDER = File($.fileName).parent;
    try { $.writeln("[MDUX] Initial setup from $.fileName: " + $.global.MDUX_JSX_FOLDER); } catch(e) {}
}

// Minimal JSON polyfill for ExtendScript environments lacking native support
if (typeof JSON === "undefined") {
    JSON = {};
}
if (typeof JSON.stringify !== "function") {
    JSON.stringify = function (value) {
        function quote(str) {
            return '"' + str.replace(/\\/g, '\\\\')
                             .replace(/"/g, '\\"')
                             .replace(/\r/g, '\\r')
                             .replace(/\n/g, '\\n')
                             .replace(/\f/g, '\\f')
                             .replace(/\t/g, '\\t')
                             .replace(/\b/g, '\\b') + '"';
        }
        function stringify(val) {
            if (val === null) return "null";
            var type = typeof val;
            if (type === "number" || type === "boolean") return String(val);
            if (type === "string") return quote(val);
            if (val instanceof Array) {
                var arr = [];
                for (var i = 0; i < val.length; i++) arr.push(stringify(val[i]));
                return "[" + arr.join(",") + "]";
            }
            if (type === "object") {
                var props = [];
                for (var key in val) {
                    if (val.hasOwnProperty(key) && typeof val[key] !== "undefined" && typeof val[key] !== "function") {
                        props.push(quote(String(key)) + ":" + stringify(val[key]));
                    }
                }
                return "{" + props.join(",") + "}";
            }
            return "null";
        }
        return stringify(value);
    };
}
if (typeof JSON.parse !== "function") {
    JSON.parse = function (text) {
        if (typeof text !== "string") return null;
        try {
            return eval('(' + text + ')');
        } catch (e) {
            return null;
        }
    };
}

function MDUX_joinPath(folderObj, filename) {
    var base = "";
    try {
        base = folderObj.fsName;
    } catch (eFs) {
        try {
            base = folderObj.toString();
        } catch (eStr) {
            base = "";
        }
    }
    if (!base) return filename;
    base = base.replace(/\\/g, "/");
    if (base.charAt(base.length - 1) === "/") {
        return base + filename;
    }
    return base + "/" + filename;
}

if (typeof MDUX === "undefined") {
    var MDUX = {};
}

// Debug logging to file
function MDUX_debugLog(message) {
    try {
        var logFile = new File(MDUX_joinPath($.global.MDUX_JSX_FOLDER || File($.fileName).parent.parent, "debug.log"));
        logFile.open("a");
        logFile.writeln("[" + new Date().toISOString() + "] " + message);
        logFile.close();
    } catch (e) {
        // Silently fail if logging doesn't work
    }
}

function MDUX_extensionRoot() {
    // Use the stored jsx folder path, or fall back to calculating it
    // from $.global.MDUX_LAST_BRIDGE_PATH (set by panel.js) or $.fileName
    if (typeof $.global.MDUX_JSX_FOLDER !== "undefined" && $.global.MDUX_JSX_FOLDER) {
        return $.global.MDUX_JSX_FOLDER;
    }
    var bridgePath = $.global.MDUX_LAST_BRIDGE_PATH || $.fileName;
    if (bridgePath) {
        var file = File(bridgePath);
        if (file.parent && file.parent.exists) {
            return file.parent;
        }
    }
    // Last resort: try to get it from $.fileName
    if ($.fileName) {
        return File($.fileName).parent;
    }
    return null;
}

function MDUX_runMagicDuctwork() {
    try {
        MDUX_debugLog("=== MDUX_runMagicDuctwork called ===");
        MDUX_debugLog("$.fileName: " + $.fileName);
        MDUX_debugLog("$.global.MDUX_LAST_BRIDGE_PATH: " + $.global.MDUX_LAST_BRIDGE_PATH);
        MDUX_debugLog("$.global.MDUX_JSX_FOLDER: " + $.global.MDUX_JSX_FOLDER);

        var root = MDUX_extensionRoot();
        if (!root) {
            MDUX_debugLog("ERROR: root is null/undefined");
            return "ERROR:Could not determine extension root folder";
        }
        var rootPath = "";
        try {
            rootPath = root.fsName || root.toString();
        } catch (e) {
            rootPath = String(root);
        }
        MDUX_debugLog("root folder: " + rootPath);
        try { $.writeln("[MDUX] root folder: " + rootPath); } catch (logRoot) {}

        var scriptFile = File(MDUX_joinPath(root, "magic-final.jsx"));
        var scriptPath = "";
        try {
            scriptPath = scriptFile.fsName || scriptFile.toString();
        } catch (e) {
            scriptPath = String(scriptFile);
        }
        MDUX_debugLog("looking for: " + scriptPath);
        MDUX_debugLog("file exists: " + scriptFile.exists);
        try { $.writeln("[MDUX] looking for: " + scriptPath); } catch (logPath) {}
        try { $.writeln("[MDUX] file exists: " + scriptFile.exists); } catch (logExists) {}

        if (!scriptFile.exists) {
            MDUX_debugLog("ERROR: magic-final missing at " + scriptPath);
            try { $.writeln("[MDUX] magic-final missing at " + scriptPath); } catch (logMiss) {}
            return "ERROR:magic-final.jsx not found at: " + scriptPath;
        }
        MDUX_debugLog("SUCCESS: About to evalFile");
        $.evalFile(scriptFile);
        MDUX_debugLog("SUCCESS: evalFile completed");
        return "OK";
    } catch (e) {
        MDUX_debugLog("EXCEPTION: " + e);
        return "ERROR:" + e;
    }
}

function MDUX_requireMagicFinal() {
    try {
        if (typeof MDUX !== "undefined" && MDUX.rotateSelection) {
            return true;
        }
        var root = MDUX_extensionRoot();
        if (!root) {
            try { $.writeln("[MDUX] requireMagicFinal: could not determine root"); } catch (logErr) {}
            return false;
        }
        var scriptFile = File(MDUX_joinPath(root, "magic-final.jsx"));
        var scriptPath = "";
        try {
            scriptPath = scriptFile.fsName || scriptFile.toString();
        } catch (e) {
            scriptPath = String(scriptFile);
        }
        try { $.writeln("[MDUX] requireMagicFinal: looking for " + scriptPath); } catch (logPath) {}
        if (!scriptFile.exists) {
            try { $.writeln("[MDUX] requireMagicFinal: file not found at " + scriptPath); } catch (logMiss) {}
            return false;
        }
        var previousForced = null;
        if ($.global.MDUX && $.global.MDUX.forcedOptions) {
            previousForced = $.global.MDUX.forcedOptions;
        }
        $.global.MDUX = $.global.MDUX || {};
        $.global.MDUX.forcedOptions = { action: "library" };
        try { $.writeln("[MDUX] requireMagicFinal: loading magic-final.jsx"); } catch (logErr) {}
        $.evalFile(scriptFile);
        try {
            if ($.global.MDUX) {
                MDUX = $.global.MDUX;
            }
        } catch (eAssign) {}
        if (previousForced) {
            $.global.MDUX.forcedOptions = previousForced;
        } else if ($.global.MDUX && $.global.MDUX.hasOwnProperty("forcedOptions")) {
            delete $.global.MDUX.forcedOptions;
        }
        var ns = (typeof MDUX !== "undefined" && MDUX.rotateSelection) ? MDUX : ($.global.MDUX || null);
        return !!(ns && (ns.rotateSelection || ns.createStandardLayerBlock || ns.importDuctworkGraphicStyles));
    } catch (e) {
        try { $.writeln("[MDUX] requireMagicFinal error: " + e); } catch (logErr2) {}
        return false;
    }
}

function MDUX_toggleIgnoreModeBridge() {
    try {
        return MDUX_toggleIgnoreMode();
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_setIgnoreModeBridge(state) {
    try {
        return MDUX_setIgnoreMode(state);
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_ignoreModeStatusBridge() {
    try {
        return MDUX_getIgnoreModeStatus();
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_skipOrthoStateBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return JSON.stringify({ available: false, error: "bridge-load" });
        }
        if (app.documents.length === 0) {
            return JSON.stringify({ available: false, reason: "no-document" });
        }
        var doc = app.activeDocument;
        var sel = null;
        try { sel = doc.selection; } catch (eSel) { sel = null; }
        if (!sel || sel.length === 0) {
            return JSON.stringify({ available: false, reason: "no-selection" });
        }
        var ns = (typeof MDUX !== "undefined" && MDUX.checkSkipOrthoState) ? MDUX : ($.global.MDUX || null);
        if (!ns || !ns.checkSkipOrthoState) {
            return "ERROR:Skip ortho state function unavailable";
        }
        var state = ns.checkSkipOrthoState(sel);
        return JSON.stringify({
            available: true,
            hasNote: !!(state && state.hasNote),
            mixed: !!(state && state.mixed)
        });
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_applyIgnoreBridge() {
    try {
        var stats = MDUX_applyIgnoreToCurrent ? MDUX_applyIgnoreToCurrent() : null;
        if (stats === null || typeof stats === "undefined") {
            stats = {};
        }
        var encodeJson = (typeof JSON !== "undefined" && JSON.stringify) ? JSON.stringify : null;
        if (!encodeJson) {
            try {
                encodeJson = function (obj) {
                    return obj ? obj.toSource() : "{}";
                };
            } catch (eSource) {
                encodeJson = function () { return "{}"; };
            }
        }
        return encodeJson(stats);
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_cleanupBridge() {
    try {
        MDUX_cleanup();
        return "OK";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_revertPreOrthoBridge() {
    try {
        if (app.documents.length === 0) {
            return "ERROR:No Illustrator document is open.";
        }
        var doc = app.activeDocument;
        var sel = null;
        try { sel = doc.selection; } catch (eSel) { sel = null; }
        if (!sel || sel.length === 0) {
            return JSON.stringify({ total: 0, reverted: 0, skipped: 0, reason: "no-selection" });
        }
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Revert function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.revertSelectionToPreOrtho) {
            var stats = MDUX.revertSelectionToPreOrtho(sel);
            return JSON.stringify(stats);
        }
        return "ERROR:Revert function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_clearRotationMetadataBridge() {
    try {
        if (app.documents.length === 0) {
            return "ERROR:No Illustrator document is open.";
        }
        var doc = app.activeDocument;
        var sel = null;
        try { sel = doc.selection; } catch (eSel) { sel = null; }
        if (!sel || sel.length === 0) {
            return "0";
        }
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Clear rotation metadata function unavailable";
        }
        if (typeof MDUX === "undefined" || !MDUX.clearRotationOverride || !MDUX.getAllPathItemsInGroup) {
            return "ERROR:Clear rotation metadata function unavailable";
        }

        var count = 0;
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];
            if (item.typename === "PathItem") {
                MDUX.clearRotationOverride(item);
                count++;
            } else if (item.typename === "GroupItem") {
                var paths = MDUX.getAllPathItemsInGroup(item);
                for (var j = 0; j < paths.length; j++) {
                    MDUX.clearRotationOverride(paths[j]);
                    count++;
                }
            }
        }

        return String(count);
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_prepareProcessBridge(optionsJSON) {
    try {
        var opts = null;
        if (optionsJSON && optionsJSON.length) {
            opts = JSON.parse(optionsJSON);
        }
        if (!opts) opts = {};
        $.global.MDUX = $.global.MDUX || {};
        $.global.MDUX.forcedOptions = {
            action: opts.action || "process",
            skipOrtho: (typeof opts.skipOrtho === "boolean") ? opts.skipOrtho : undefined,
            rotationOverride: (typeof opts.rotationOverride === "number" && isFinite(opts.rotationOverride)) ? opts.rotationOverride : null,
            skipAllBranchSegments: !!opts.skipAllBranchSegments,
            skipFinalRegisterSegment: !!opts.skipFinalRegisterSegment
        };
        return "OK";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_rotateSelectionBridge(angle) {
    try {
        if (typeof angle !== "number") angle = parseFloat(angle);
        if (!isFinite(angle)) return "ERROR:Invalid rotation value";
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Rotate function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.rotateSelection) {
            var stats = MDUX.rotateSelection(angle);
            return JSON.stringify(stats);
        }
        return "ERROR:Rotate function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_scaleSelectionBridge(percent) {
    try {
        if (typeof percent !== "number") percent = parseFloat(percent);
        if (!isFinite(percent)) return "ERROR:Invalid scale value";
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Scale function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.scaleSelectionAbsolute) {
            var stats = MDUX.scaleSelectionAbsolute(percent);
            return JSON.stringify(stats);
        }
        return "ERROR:Scale function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_resetScaleBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Reset scale function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.resetSelectionScale) {
            var stats = MDUX.resetSelectionScale();
            return JSON.stringify(stats);
        }
        return "ERROR:Reset scale function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_rotationStateBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Rotation function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.getRotationOverrideSummary) {
            var summary = MDUX.getRotationOverrideSummary();
            return JSON.stringify(summary);
        }
        return "ERROR:Rotation function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_isolatePartsBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Isolate parts function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.isolateDuctworkParts) {
            return MDUX.isolateDuctworkParts();
        }
        return "ERROR:Isolate parts function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_isolateDuctworkBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Isolate ductwork function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.isolateDuctworkLines) {
            return MDUX.isolateDuctworkLines();
        }
        return "ERROR:Isolate ductwork function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_unlockDuctworkBridge() {
    try {
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Unlock function unavailable (bridge load failed)";
        }
        var ns = (typeof MDUX !== "undefined" && MDUX.unlockAllDuctworkLayers) ? MDUX : ($.global.MDUX || null);
        if (ns && ns.unlockAllDuctworkLayers) {
            return ns.unlockAllDuctworkLayers();
        }
        return "ERROR:Unlock function unavailable (namespace missing)";
    } catch (e) {
        return "ERROR:Unlock (exception): " + e;
    }
}

function MDUX_importGraphicStylesBridge() {
    try {
        if (app.documents.length === 0) {
            return "ERROR:No Illustrator document is open.";
        }
        var destDoc = app.activeDocument;
        var sourceFile = new File("E:/Work/Work/Floorplans/Ductwork Assets/DuctworkLines.ai");
        if (!sourceFile.exists) {
            return "ERROR:DuctworkLines.ai not found.";
        }

        var sourceDoc = null;
        try { $.writeln("[MDUX] Import styles: opening " + sourceFile.fsName); } catch (logOpen) {}
        try {
            sourceDoc = app.open(sourceFile);
            app.activeDocument = sourceDoc;
            try {
                for (var L = 0; L < sourceDoc.layers.length; L++) {
                    try { sourceDoc.layers[L].locked = false; } catch (eLock) {}
                    try { sourceDoc.layers[L].visible = true; } catch (eVis) {}
                }
            } catch (eIter) {}

            var items = null;
            try { items = sourceDoc.pageItems; } catch (eItems) { items = null; }
            if (!items || items.length === 0) {
                sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
                app.activeDocument = destDoc;
                return "ERROR:Source document contained no artwork.";
            }

            for (var i = 0; i < items.length; i++) {
                try { items[i].selected = true; } catch (eSel) {}
            }
            app.copy();
            sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (copyErr) {
            try {
                if (sourceDoc) sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeErr) {}
            app.activeDocument = destDoc;
            try { $.writeln("[MDUX] Import styles error: " + copyErr); } catch (logCopy) {}
            return "ERROR:" + copyErr;
        }

        app.activeDocument = destDoc;
        destDoc.selection = null;
        var tempLayerName = "__MDUX_STYLE_IMPORT__";
        var tempLayer = null;
        try {
            tempLayer = destDoc.layers.getByName(tempLayerName);
        } catch (findErr) {
            tempLayer = destDoc.layers.add();
            tempLayer.name = tempLayerName;
        }
        try { tempLayer.locked = false; } catch (lockErr) {}
        destDoc.activeLayer = tempLayer;
        app.executeMenuCommand("pasteInPlace");
        var pasted = null;
        try { pasted = destDoc.selection; } catch (selErr) { pasted = null; }
        if (pasted) {
            if (pasted.length === undefined) pasted = [pasted];
            for (var p = pasted.length - 1; p >= 0; p--) {
                try { pasted[p].remove(); } catch (remErr) {}
            }
        }
        destDoc.selection = null;
        try {
            if (!tempLayer.pageItems.length && !tempLayer.groupItems.length) {
                tempLayer.remove();
            }
        } catch (cleanupErr) {}
        try { app.redraw(); } catch (redErr) {}
        try { $.writeln("[MDUX] Import styles completed successfully"); } catch (logDone) {}
        return "Graphic styles imported.";
    } catch (e) {
        try { $.writeln("[MDUX] Import styles exception: " + e); } catch (logFinal) {}
        return "ERROR:Import styles (exception): " + e;
    }
}

function MDUX_createLayersBridge() {
    try {
        if (app.documents.length === 0) {
            return "ERROR:No Illustrator document is open.";
        }
        var doc = app.activeDocument;
        var desired = [
            "Scale Factor Container Layer",
            "Frame",
            "Ignored",
            "Thermostats",
            "Units",
            "Secondary Exhaust Registers",
            "Thermostat Lines",
            "Exhaust Registers",
            "Rectangular Registers",
            "Circular Registers",
            "Orange Register",
            "Square Registers",
            "Light Orange Ductwork",
            "Orange Ductwork",
            "Blue Ductwork",
            "Green Ductwork"
        ];

    for (var i = 0; i < desired.length; i++) {
        var name = desired[i];
        var layer = null;
        try { layer = doc.layers.getByName(name); } catch (eFind) {
            try {
                layer = doc.layers.add();
                layer.name = name;
                try { $.writeln("[MDUX] Created missing layer: " + name); } catch (logCreate) {}
            } catch (eCreate) {
                layer = null;
            }
        }
    }

    for (var idx = desired.length - 1; idx >= 0; idx--) {
        try {
            var target = doc.layers.getByName(desired[idx]);
            if (!target) continue;
            var prevLocked = null;
            try { prevLocked = target.locked; target.locked = false; } catch (eLock) {}
            try { target.visible = true; } catch (eVisible) {}
            try {
                target.move(doc.layers[0], ElementPlacement.PLACEBEFORE);
            } catch (eMove) {}
            try {
                if (prevLocked !== null) target.locked = prevLocked;
            } catch (eRestore) {}
        } catch (eTarget) {}
    }

    try { app.redraw(); } catch (eRedraw) {}
    try { $.writeln("[MDUX] Ensure standard layers completed"); } catch (logDone) {}
    return "Standard ductwork layers ensured.";
} catch (e) {
    try { $.writeln("[MDUX] Create layers exception: " + e); } catch (logErr) {}
    return "ERROR:Create layers (exception): " + e;
}
}

function MDUX_runEmoryDuctwork(createRegisterWires) {
    try {
        if (!app.documents.length) {
            return "ERROR:No document open";
        }

        var doc = app.activeDocument;
        var REGISTER_WIRE_TAG = "MD:REGISTER_WIRE";
        var WIRE_CONNECTION_TOLERANCE = 50; // Increased tolerance to 50px
        var swappedCount = 0;
        var wireCount = 0;

        // Default to false if not specified
        if (typeof createRegisterWires === 'undefined') {
            createRegisterWires = false;
        }

        $.writeln("[EMORY] Starting Emory ductwork processing");

        // Step 1: Swap all placed items to Emory versions
        var allItems = doc.pageItems;
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            if (item.typename === "PlacedItem") {
                try {
                    var currentFile = item.file;
                    if (currentFile) {
                        var currentPath = currentFile.fsName;
                        if (currentPath.indexOf(" Emory") === -1) {
                            var emoryPath = currentPath.replace(/(\.[^.]+)$/, " Emory$1");
                            var emoryFile = new File(emoryPath);
                            if (emoryFile.exists) {
                                item.file = emoryFile;
                                swappedCount++;
                            }
                        }
                    }
                } catch (e) {}
            }
        }

        // Step 1.5: Place/swap register images at anchors based on layer
        $.writeln("[EMORY] Placing register images at anchors based on layer");
        var registerLayerMap = {
            "Square Registers": "Square Register Emory",
            "Rectangular Registers": "Rectangular Register Emory"
        };

        for (var layerName in registerLayerMap) {
            var targetImageName = registerLayerMap[layerName];
            var registerLayer = null;

            // Find the layer
            try {
                for (var li = 0; li < doc.layers.length; li++) {
                    if (doc.layers[li].name === layerName) {
                        registerLayer = doc.layers[li];
                        break;
                    }
                }
            } catch (e) {}

            if (!registerLayer || !registerLayer.pathItems) continue;

            // Get all anchor points on this layer
            for (var pi = 0; pi < registerLayer.pathItems.length; pi++) {
                var regPath = registerLayer.pathItems[pi];
                if (!regPath || !regPath.pathPoints || regPath.pathPoints.length === 0) continue;

                var anchor = regPath.pathPoints[0].anchor;
                var anchorPos = { x: anchor[0], y: anchor[1] };

                // Check if there's already a PlacedItem near this anchor
                var foundPlacedItem = null;
                var searchTolerance = 5; // 5px tolerance

                for (var ai = 0; ai < allItems.length; ai++) {
                    var checkItem = allItems[ai];
                    if (checkItem.typename !== "PlacedItem") continue;

                    try {
                        var itemPos = { x: checkItem.position[0], y: checkItem.position[1] };
                        var dx = itemPos.x - anchorPos.x;
                        var dy = itemPos.y - anchorPos.y;
                        var dist = Math.sqrt(dx * dx + dy * dy);

                        if (dist < searchTolerance) {
                            foundPlacedItem = checkItem;
                            break;
                        }
                    } catch (e) {}
                }

                // Swap or place the correct Emory register image
                try {
                    // Find the Emory register file based on layer
                    var basePath = "E:\\Work\\Work\\Floorplans\\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\\Illustrator\\";
                    var emoryRegisterPath = basePath + "Ductwork Pieces Emory\\" + targetImageName + ".ai";
                    var emoryRegisterFile = new File(emoryRegisterPath);

                    if (!emoryRegisterFile.exists) {
                        // Try alternate path
                        emoryRegisterPath = basePath + "Ductwork Pieces\\" + targetImageName + ".ai";
                        emoryRegisterFile = new File(emoryRegisterPath);
                    }

                    if (emoryRegisterFile.exists) {
                        if (foundPlacedItem) {
                            // Swap existing placed item to correct Emory register
                            var currentFileName = foundPlacedItem.file ? foundPlacedItem.file.name : "";
                            if (currentFileName !== targetImageName + ".ai") {
                                foundPlacedItem.file = emoryRegisterFile;
                                swappedCount++;
                                $.writeln("[EMORY] Swapped register at [" + anchorPos.x.toFixed(1) + "," + anchorPos.y.toFixed(1) + "] to " + targetImageName);
                            }
                        } else {
                            // Place new Emory register at anchor
                            var newPlaced = registerLayer.placedItems.add();
                            newPlaced.file = emoryRegisterFile;
                            newPlaced.position = [anchorPos.x, anchorPos.y];
                            swappedCount++;
                            $.writeln("[EMORY] Placed " + targetImageName + " at [" + anchorPos.x.toFixed(1) + "," + anchorPos.y.toFixed(1) + "]");
                        }
                    } else {
                        $.writeln("[EMORY] Could not find Emory register file: " + emoryRegisterPath);
                    }
                } catch (e) {
                    $.writeln("[EMORY] Error placing/swapping register: " + e);
                }
            }
        }

        // Step 2: Generate register wires (only if enabled)
        if (createRegisterWires) {
            $.writeln("[EMORY] Register wire creation enabled");
            var ductworkPaths = [];
            var registerPoints = [];
            var ignoreAnchors = [];

            $.writeln("[EMORY] Scanning for ductwork paths, registers, and ignore anchors");

        // Find ductwork paths
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            if (item.typename === "PathItem") {
                var layerName = "";
                try { layerName = item.layer.name; } catch (e) {}
                if (layerName && (layerName.indexOf("Ductwork") !== -1 || layerName.indexOf("ductwork") !== -1)) {
                    ductworkPaths.push(item);
                }
            }
        }

        // Find register anchor points on register layers
        var registerLayerNames = ["Square Registers", "Rectangular Registers", "Registers"];
        for (var layerIdx = 0; layerIdx < registerLayerNames.length; layerIdx++) {
            var layerName = registerLayerNames[layerIdx];
            var registerLayer = null;
            try {
                for (var li = 0; li < doc.layers.length; li++) {
                    if (doc.layers[li].name === layerName) {
                        registerLayer = doc.layers[li];
                        break;
                    }
                }
            } catch (e) {}

            if (registerLayer && registerLayer.pathItems) {
                for (var pi = 0; pi < registerLayer.pathItems.length; pi++) {
                    var regPath = registerLayer.pathItems[pi];
                    if (regPath && regPath.pathPoints && regPath.pathPoints.length > 0) {
                        // Use first anchor point as register location
                        var anchor = regPath.pathPoints[0].anchor;
                        registerPoints.push({ x: anchor[0], y: anchor[1] });
                    }
                }
            }
        }

        // Find ignore anchors on the ignore layer
        var ignoreLayer = null;
        try {
            for (var li = 0; li < doc.layers.length; li++) {
                if (doc.layers[li].name === "Ignore" || doc.layers[li].name === "ignore") {
                    ignoreLayer = doc.layers[li];
                    break;
                }
            }
        } catch (e) {}

        if (ignoreLayer && ignoreLayer.pathItems) {
            for (var pi = 0; pi < ignoreLayer.pathItems.length; pi++) {
                var ignorePath = ignoreLayer.pathItems[pi];
                if (ignorePath && ignorePath.pathPoints) {
                    for (var pp = 0; pp < ignorePath.pathPoints.length; pp++) {
                        var anchor = ignorePath.pathPoints[pp].anchor;
                        ignoreAnchors.push({ x: anchor[0], y: anchor[1] });
                    }
                }
            }
        }

        $.writeln("[EMORY] Found " + ductworkPaths.length + " ductwork paths");
        $.writeln("[EMORY] Found " + registerPoints.length + " register points");
        $.writeln("[EMORY] Found " + ignoreAnchors.length + " ignore anchors");

        for (var i = 0; i < ductworkPaths.length; i++) {
            var path = ductworkPaths[i];
            if (!path.pathPoints || path.pathPoints.length < 2) continue; // Need at least 2 points
            var layer = path.layer;
            if (!layer) continue;

            // Check both ends of the ductwork path
            var endpoints = [
                { endIdx: 0, prevIdx: 1, name: "start" },
                { endIdx: path.pathPoints.length - 1, prevIdx: path.pathPoints.length - 2, name: "end" }
            ];

            for (var j = 0; j < endpoints.length; j++) {
                var epInfo = endpoints[j];
                if (epInfo.prevIdx >= path.pathPoints.length || epInfo.prevIdx < 0) continue;

                var endAnchor = path.pathPoints[epInfo.endIdx].anchor;
                var prevAnchor = path.pathPoints[epInfo.prevIdx].anchor;
                var endPoint = { x: endAnchor[0], y: endAnchor[1] };
                var prevPoint = { x: prevAnchor[0], y: prevAnchor[1] };

                // Check if endpoint is near an ignore anchor (if so, SKIP it)
                var nearIgnore = false;
                for (var k = 0; k < ignoreAnchors.length; k++) {
                    var ignoreAnchor = ignoreAnchors[k];
                    var dx = ignoreAnchor.x - endPoint.x;
                    var dy = ignoreAnchor.y - endPoint.y;
                    var dist = Math.sqrt(dx * dx + dy * dy);
                    if (dist < 5) { // Within 5px of ignore anchor
                        nearIgnore = true;
                        $.writeln("[EMORY] Skipping endpoint at [" + endPoint.x.toFixed(1) + "," + endPoint.y.toFixed(1) + "] - near ignore anchor (dist=" + dist.toFixed(1) + ")");
                        break;
                    }
                }

                if (nearIgnore) continue;

                // Check if endpoint is near a register
                var nearRegister = false;
                for (var k = 0; k < registerPoints.length; k++) {
                    var register = registerPoints[k];
                    var dx = register.x - endPoint.x;
                    var dy = register.y - endPoint.y;
                    var dist = Math.sqrt(dx * dx + dy * dy);
                    if (dist < WIRE_CONNECTION_TOLERANCE) {
                        nearRegister = true;
                        $.writeln("[EMORY] Endpoint at [" + endPoint.x.toFixed(1) + "," + endPoint.y.toFixed(1) + "] is near register at [" + register.x.toFixed(1) + "," + register.y.toFixed(1) + "] (dist=" + dist.toFixed(1) + ")");
                        break;
                    }
                }

                if (nearRegister) {
                    try {
                        $.writeln("[EMORY] ===== WIRE CREATION DEBUG =====");
                        $.writeln("[EMORY] Path has " + path.pathPoints.length + " points");
                        $.writeln("[EMORY] Endpoint info: name=" + epInfo.name + ", endIdx=" + epInfo.endIdx + ", prevIdx=" + epInfo.prevIdx);
                        $.writeln("[EMORY] prevPoint: [" + prevPoint.x.toFixed(1) + "," + prevPoint.y.toFixed(1) + "]");
                        $.writeln("[EMORY] endPoint: [" + endPoint.x.toFixed(1) + "," + endPoint.y.toFixed(1) + "]");

                        // Create wire from previous point to endpoint (last segment of ductwork)
                        var wireDX = endPoint.x - prevPoint.x;
                        var wireDY = endPoint.y - prevPoint.y;
                        var wireLen = Math.sqrt(wireDX * wireDX + wireDY * wireDY);

                        $.writeln("[EMORY] Wire vector: dx=" + wireDX.toFixed(2) + ", dy=" + wireDY.toFixed(2) + ", length=" + wireLen.toFixed(2));

                        if (wireLen < 5) {
                            $.writeln("[EMORY] Skipping wire - too short (< 5px): " + wireLen.toFixed(2));
                            continue;
                        }

                        // Create wire path from previous point to endpoint
                        var wirePath = layer.pathItems.add();
                        wirePath.setEntirePath([[prevPoint.x, prevPoint.y], [endPoint.x, endPoint.y]]);

                        $.writeln("[EMORY] Wire path created, checking points...");
                        if (wirePath.pathPoints && wirePath.pathPoints.length === 2) {
                            var wp0 = wirePath.pathPoints[0].anchor;
                            var wp1 = wirePath.pathPoints[1].anchor;
                            $.writeln("[EMORY] Wire point 0: [" + wp0[0].toFixed(1) + "," + wp0[1].toFixed(1) + "]");
                            $.writeln("[EMORY] Wire point 1: [" + wp1[0].toFixed(1) + "," + wp1[1].toFixed(1) + "]");
                            var actualDX = wp1[0] - wp0[0];
                            var actualDY = wp1[1] - wp0[1];
                            var actualLen = Math.sqrt(actualDX * actualDX + actualDY * actualDY);
                            $.writeln("[EMORY] Actual wire length: " + actualLen.toFixed(2) + "px");
                        } else {
                            $.writeln("[EMORY] ERROR: Wire has " + (wirePath.pathPoints ? wirePath.pathPoints.length : "NO") + " points!");
                        }

                        // Style the wire: stroke-only (explicitly clear fills) - from Emory script
                        wirePath.closed = false;

                        // Explicitly remove any graphic style that might be applied
                        try {
                            wirePath.unapplyAll();
                        } catch(e) {}

                        wirePath.stroked = true;
                        wirePath.strokeWidth = 3;
                        try {
                            var wireColor = new RGBColor();
                            wireColor.red = 0;
                            wireColor.green = 0;
                            wireColor.blue = 255;
                            wirePath.strokeColor = wireColor;
                        } catch (eWireColor) {}
                        wirePath.strokeCap = StrokeCap.ROUNDENDCAP;
                        wirePath.strokeJoin = StrokeJoin.ROUNDENDJOIN;

                        wirePath.filled = false;
                        try {
                            var noColor = new NoColor();
                            wirePath.fillColor = noColor;
                        } catch (e) {}

                        // Add curved handles - using Emory script logic
                        var wirePoints = wirePath.pathPoints;
                        if (wirePoints && wirePoints.length === 2) {
                            // Get segment direction (from previous segment if possible)
                            var segDirection = { x: wireDX / wireLen, y: wireDY / wireLen };

                            // Try to get the previous segment's direction for smoother curve
                            if (path.pathPoints.length > 2 && epInfo.prevIdx > 0 && epInfo.prevIdx < path.pathPoints.length - 1) {
                                var prevPrevAnchor = path.pathPoints[epInfo.prevIdx - (epInfo.endIdx > epInfo.prevIdx ? -1 : 1)].anchor;
                                var prevPrevPoint = { x: prevPrevAnchor[0], y: prevPrevAnchor[1] };
                                var prevSegDX = prevPoint.x - prevPrevPoint.x;
                                var prevSegDY = prevPoint.y - prevPrevPoint.y;
                                var prevSegLen = Math.sqrt(prevSegDX * prevSegDX + prevSegDY * prevSegDY);
                                if (prevSegLen > 0) {
                                    segDirection = { x: prevSegDX / prevSegLen, y: prevSegDY / prevSegLen };
                                }
                            }

                            // Scale handles based on wire length - from Emory script
                            var handleLen;
                            if (wireLen < 10) {
                                handleLen = wireLen * 0.05;
                            } else if (wireLen < 20) {
                                var t = (wireLen - 10) / 10;
                                handleLen = wireLen * (0.05 + t * 0.05);
                            } else if (wireLen < 30) {
                                var t = (wireLen - 20) / 10;
                                handleLen = wireLen * (0.10 + t * 0.05);
                            } else if (wireLen < 50) {
                                var t = (wireLen - 30) / 20;
                                handleLen = wireLen * (0.15 + t * 0.10);
                            } else {
                                handleLen = Math.min(wireLen * 0.30, 30);
                            }
                            handleLen = Math.max(handleLen, 0.5);

                            var startPointInfo = wirePoints[0];
                            var endPointInfo = wirePoints[1];

                            // Start handle: extend along the previous segment direction
                            if (startPointInfo) {
                                startPointInfo.leftDirection = [prevPoint.x, prevPoint.y];
                                startPointInfo.rightDirection = [
                                    prevPoint.x + segDirection.x * handleLen,
                                    prevPoint.y + segDirection.y * handleLen
                                ];
                                try { startPointInfo.pointType = PointType.SMOOTH; } catch (eWireStartType) {}
                            }

                            // End handle: point downward (inverted)
                            if (endPointInfo) {
                                endPointInfo.leftDirection = [
                                    endPoint.x,
                                    endPoint.y + handleLen * 0.6
                                ];
                                endPointInfo.rightDirection = [endPoint.x, endPoint.y];
                                try { endPointInfo.pointType = PointType.SMOOTH; } catch (eWireEndType) {}
                            }
                        }

                        wirePath.name = "Register Wire";
                        wirePath.note = REGISTER_WIRE_TAG;
                        wireCount++;
                        $.writeln("[EMORY] Wire created successfully");
                    } catch (eWire) {
                        $.writeln("[EMORY] Wire creation error: " + eWire);
                    }
                }
            }
        }
        } else {
            $.writeln("[EMORY] Register wire creation skipped (disabled)");
        }

        var message = "Swapped " + swappedCount + " items to Emory versions";
        if (createRegisterWires) {
            message += " and created " + wireCount + " register wires";
        }
        message += ".";

        $.writeln("[EMORY] Complete: " + message);
        return message;
    } catch (e) {
        $.writeln("[EMORY] Error: " + e);
        return "ERROR:" + e.toString();
    }
}

function MDUX_moveToLayerBridge(optionsJSON) {
    try {
        $.writeln("[MOVE] MDUX_moveToLayerBridge called");
        $.writeln("[MOVE] Input JSON: " + optionsJSON);

        if (!app.documents.length) {
            $.writeln("[MOVE] ERROR: No documents open");
            return JSON.stringify({ itemsMoved: 0, anchorsMoved: 0, reason: 'no-document' });
        }

        var doc = app.activeDocument;

        // Clean up any leftover temp layers from previous errors
        try {
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layerName = doc.layers[i].name;
                if (layerName.indexOf('MDUX_TEMP_') === 0 || layerName === 'Scale Factor Container Layer') {
                    $.writeln("[MOVE] Cleaning up temp layer: " + layerName);
                    doc.layers[i].remove();
                }
            }
        } catch (e) {
            $.writeln("[MOVE] Error cleaning temp layers: " + e);
        }

        var selection = doc.selection;
        $.writeln("[MOVE] Selection length: " + (selection ? selection.length : 0));

        if (!selection || selection.length === 0) {
            $.writeln("[MOVE] ERROR: No selection");
            return JSON.stringify({ itemsMoved: 0, anchorsMoved: 0, reason: 'no-selection' });
        }

        var options = JSON.parse(optionsJSON);
        var targetLayerName = options.layerName;
        var fileBaseName = options.fileBaseName;
        $.writeln("[MOVE] Target layer: " + targetLayerName);
        $.writeln("[MOVE] File base name: " + fileBaseName);

        // Get or create target layer
        var targetLayer = null;
        try {
            targetLayer = doc.layers.getByName(targetLayerName);
            $.writeln("[MOVE] Found existing layer: " + targetLayerName);
        } catch (e) {
            targetLayer = doc.layers.add();
            targetLayer.name = targetLayerName;
            $.writeln("[MOVE] Created new layer: " + targetLayerName);
        }

        var isIgnoreLayer = (targetLayerName === 'Ignore' || targetLayerName === 'Ignored');
        var wasLocked = targetLayer.locked;
        var wasVisible = targetLayer.visible;
        $.writeln("[MOVE] Layer locked: " + wasLocked + ", visible: " + wasVisible);

        // Unlock and show target layer temporarily
        if (isIgnoreLayer || targetLayer.locked) {
            targetLayer.locked = false;
            $.writeln("[MOVE] Unlocked target layer");
        }
        if (!targetLayer.visible) {
            targetLayer.visible = true;
            $.writeln("[MOVE] Made target layer visible");
        }

        var itemsMoved = 0;
        var anchorsMoved = 0;
        var itemsSkipped = 0;
        var filePath = null;

        if (fileBaseName) {
            var COMPONENT_FILES_PATH = 'E:/Work/Work/Floorplans/Ductwork Assets/';
            filePath = new File(COMPONENT_FILES_PATH + fileBaseName);
            $.writeln("[MOVE] File path: " + filePath.fsName);
            if (!filePath.exists) {
                $.writeln("[MOVE] WARNING: File does not exist");
                filePath = null;
            } else {
                $.writeln("[MOVE] File exists");
            }
        }

        // Define valid ductwork parts layers
        var validDuctworkLayers = [
            'Units',
            'Square Registers',
            'Rectangular Registers',
            'Circular Registers',
            'Exhaust Registers',
            'Secondary Exhaust Registers',
            'Thermostats',
            'Ignore',
            'Ignored'
        ];

        // Helper function to check if layer is valid
        function isValidDuctworkLayer(layerName) {
            for (var i = 0; i < validDuctworkLayers.length; i++) {
                if (validDuctworkLayers[i] === layerName) {
                    return true;
                }
            }
            return false;
        }

        // Helper function to get scale of a PlacedItem using matrix
        function getItemScale(item) {
            try {
                if (item.typename !== 'PlacedItem') return 100;

                // Get the transformation matrix
                var matrix = item.matrix;

                // Calculate scale from matrix
                // mValueA is horizontal scale, mValueD is vertical scale
                var scaleX = Math.sqrt(matrix.mValueA * matrix.mValueA + matrix.mValueB * matrix.mValueB);
                var scaleY = Math.sqrt(matrix.mValueC * matrix.mValueC + matrix.mValueD * matrix.mValueD);

                // Use average of X and Y scale
                var scale = (scaleX + scaleY) / 2;

                return scale * 100;
            } catch (e) {
                $.writeln("[MOVE] Error getting scale: " + e);
                return 100;
            }
        }

        // Find smallest scale on target layer (if not Ignore layer)
        var smallestScale = 100;
        if (!isIgnoreLayer && filePath) {
            $.writeln("[MOVE] Scanning target layer for smallest scale...");
            try {
                var layerItems = targetLayer.pageItems;
                for (var i = 0; i < layerItems.length; i++) {
                    if (layerItems[i].typename === 'PlacedItem') {
                        var scale = getItemScale(layerItems[i]);
                        if (scale < smallestScale) {
                            smallestScale = scale;
                        }
                    }
                }
                $.writeln("[MOVE] Smallest scale found: " + smallestScale + "%");
            } catch (e) {
                $.writeln("[MOVE] Error scanning for scale: " + e);
                smallestScale = 100;
            }
        }

        // Build map of anchor positions (PathItems with 1 point)
        var anchors = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === 'PathItem' && selection[i].pathPoints.length === 1) {
                var anchor = selection[i].pathPoints[0].anchor;
                anchors.push({ x: anchor[0], y: anchor[1] });
            }
        }
        $.writeln("[MOVE] Found " + anchors.length + " anchors in selection");

        // Helper function to find anchors at PlacedItem center
        function findAnchorsAtCenter(centerX, centerY, tolerance) {
            var foundAnchors = [];
            var searchLog = [];
            $.writeln("[MOVE] Searching for anchors at center [" + centerX + ", " + centerY + "] with tolerance " + tolerance);
            searchLog.push("Searching for anchors at center [" + centerX.toFixed(2) + ", " + centerY.toFixed(2) + "] with tolerance " + tolerance);

            // Search all ductwork parts layers for anchors
            for (var layerIdx = 0; layerIdx < doc.layers.length; layerIdx++) {
                var layer = doc.layers[layerIdx];
                if (!isValidDuctworkLayer(layer.name)) continue;

                searchLog.push("Checking layer: " + layer.name + " (" + layer.pathItems.length + " paths)");

                try {
                    // Check all PathItems on this layer
                    for (var pathIdx = 0; pathIdx < layer.pathItems.length; pathIdx++) {
                        var path = layer.pathItems[pathIdx];

                        // Only consider 1-point paths (anchors)
                        if (path.pathPoints && path.pathPoints.length === 1) {
                            var anchorPt = path.pathPoints[0].anchor;
                            var dx = anchorPt[0] - centerX;
                            var dy = anchorPt[1] - centerY;
                            var dist = Math.sqrt(dx * dx + dy * dy);

                            searchLog.push("  Found 1-pt anchor at [" + anchorPt[0].toFixed(2) + ", " + anchorPt[1].toFixed(2) + "] - distance: " + dist.toFixed(2) + "px");

                            if (dist <= tolerance) {
                                $.writeln("[MOVE]   Found anchor at [" + anchorPt[0] + ", " + anchorPt[1] + "] on layer '" + layer.name + "' (distance: " + dist.toFixed(2) + ")");
                                searchLog.push("    MATCH! Adding to foundAnchors");
                                foundAnchors.push(path);
                            }
                        }
                    }
                } catch (eLayerSearch) {
                    $.writeln("[MOVE]   Error searching layer '" + layer.name + "': " + eLayerSearch);
                    searchLog.push("  Error: " + eLayerSearch);
                }
            }

            $.writeln("[MOVE] Found " + foundAnchors.length + " anchors at center");
            searchLog.push("Found " + foundAnchors.length + " anchors at center");

            // Write search log to file
            try {
                var debugLogPath = "C:/Users/Chris/AppData/Roaming/Adobe/CEP/extensions/Magic-Ductwork-Panel/anchor-search-debug.log";
                var debugFile = new File(debugLogPath);
                if (debugFile.open("w")) {
                    for (var i = 0; i < searchLog.length; i++) {
                        debugFile.writeln(searchLog[i]);
                    }
                    debugFile.close();
                }
            } catch (eDebug) {
                $.writeln("[MOVE] Error writing anchor search debug log: " + eDebug);
            }

            return foundAnchors;
        }

        // Move selection to target layer
        $.writeln("[MOVE] Processing " + selection.length + " selected items...");
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            $.writeln("[MOVE] Item " + i + " typename: " + item.typename);

            // Check if item is on a valid ductwork parts layer
            var itemLayerName = item.layer ? item.layer.name : null;
            $.writeln("[MOVE]   Item layer: " + itemLayerName);

            if (!isValidDuctworkLayer(itemLayerName)) {
                $.writeln("[MOVE]   SKIPPED - not on a ductwork parts layer");
                itemsSkipped++;
                continue;
            }

            try {
                // Move PlacedItems and replace with fresh file centered on anchor
                if (item.typename === 'PlacedItem') {
                    $.writeln("[MOVE]   Processing PlacedItem...");

                    // If we have a file to replace with, reload centered on anchor with smallest scale
                    if (filePath && !isIgnoreLayer) {
                        try {
                            $.writeln("[MOVE]   Replacing with fresh file...");

                            // Get item's geometric center
                            var bounds = item.geometricBounds;
                            var centerX = (bounds[0] + bounds[2]) / 2;
                            var centerY = (bounds[1] + bounds[3]) / 2;
                            $.writeln("[MOVE]   Item center: " + centerX + ", " + centerY);

                            // Find nearest anchor
                            var nearestAnchor = null;
                            var minDist = 999999;
                            for (var a = 0; a < anchors.length; a++) {
                                var dx = anchors[a].x - centerX;
                                var dy = anchors[a].y - centerY;
                                var dist = Math.sqrt(dx * dx + dy * dy);
                                if (dist < minDist) {
                                    minDist = dist;
                                    nearestAnchor = anchors[a];
                                }
                            }

                            // Use anchor position if found, otherwise use item center
                            var targetX = nearestAnchor ? nearestAnchor.x : centerX;
                            var targetY = nearestAnchor ? nearestAnchor.y : centerY;
                            $.writeln("[MOVE]   Target center (anchor): " + targetX + ", " + targetY);

                            // Delete the old item
                            item.remove();
                            $.writeln("[MOVE]   Old item removed");

                            // Place fresh item from file
                            var newItem = targetLayer.placedItems.add();
                            newItem.file = filePath;

                            // Get new item bounds to calculate center offset
                            var newBounds = newItem.geometricBounds;
                            var newWidth = Math.abs(newBounds[2] - newBounds[0]);
                            var newHeight = Math.abs(newBounds[1] - newBounds[3]);

                            // Scale FIRST if needed
                            if (smallestScale !== 100) {
                                newItem.resize(smallestScale, smallestScale, true, false, false, false, 100, Transformation.CENTER);
                                $.writeln("[MOVE]   Scaled to " + smallestScale + "%");

                                // Recalculate bounds after scaling
                                newBounds = newItem.geometricBounds;
                                newWidth = Math.abs(newBounds[2] - newBounds[0]);
                                newHeight = Math.abs(newBounds[1] - newBounds[3]);
                            }

                            // Position so center aligns with target anchor
                            newItem.position = [targetX - newWidth / 2, targetY + newHeight / 2];
                            $.writeln("[MOVE]   Item centered on anchor at " + targetX + ", " + targetY);

                            itemsMoved++;

                            // Find and move any anchors at this location (5px tolerance)
                            try {
                                var replacementAnchors = findAnchorsAtCenter(targetX, targetY, 5);
                                for (var anchorIdx = 0; anchorIdx < replacementAnchors.length; anchorIdx++) {
                                    try {
                                        var anchorLayer = replacementAnchors[anchorIdx].layer ? replacementAnchors[anchorIdx].layer.name : null;
                                        if (anchorLayer === targetLayerName) {
                                            $.writeln("[MOVE]     Anchor already on target layer, skipping");
                                            continue;
                                        }

                                        var wasInSelection = false;
                                        for (var selIdx = 0; selIdx < selection.length; selIdx++) {
                                            if (selection[selIdx] === replacementAnchors[anchorIdx]) {
                                                wasInSelection = true;
                                                break;
                                            }
                                        }

                                        if (!wasInSelection) {
                                            $.writeln("[MOVE]     Moving anchor from '" + anchorLayer + "' to '" + targetLayerName + "'");
                                            replacementAnchors[anchorIdx].move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                                            anchorsMoved++;
                                        }
                                    } catch (eAnchor) {
                                        $.writeln("[MOVE]     Error moving anchor: " + eAnchor);
                                    }
                                }
                            } catch (eSearchAnchors) {
                                $.writeln("[MOVE]   Error searching for anchors after replacement: " + eSearchAnchors);
                            }
                        } catch (eReplace) {
                            $.writeln("[MOVE]   ERROR replacing item: " + eReplace);
                            // If replacement failed and item still exists, just move it
                            try {
                                if (item && !item.removed) {
                                    item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                                    itemsMoved++;
                                }
                            } catch (e) {}
                        }
                    } else {
                        // No file to replace with, just move the item
                        $.writeln("[MOVE]   Moving PlacedItem without replacement...");

                        // Get item center BEFORE moving it
                        var bounds = item.geometricBounds;
                        var centerX = (bounds[0] + bounds[2]) / 2;
                        var centerY = (bounds[1] + bounds[3]) / 2;

                        item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                        itemsMoved++;
                        $.writeln("[MOVE]   PlacedItem moved successfully");

                        // Find and move any anchors at this item's center (10px tolerance - increased from 5px)
                        $.writeln("[MOVE]   Searching for anchors at center [" + centerX.toFixed(2) + ", " + centerY.toFixed(2) + "]");
                        var centerAnchors = findAnchorsAtCenter(centerX, centerY, 10);
                        $.writeln("[MOVE]   Found " + centerAnchors.length + " anchors at center");
                        for (var anchorIdx = 0; anchorIdx < centerAnchors.length; anchorIdx++) {
                            try {
                                // Skip if anchor is already on target layer or was in the selection
                                var anchorLayer = centerAnchors[anchorIdx].layer ? centerAnchors[anchorIdx].layer.name : null;
                                if (anchorLayer === targetLayerName) {
                                    $.writeln("[MOVE]     Anchor already on target layer, skipping");
                                    continue;
                                }

                                // Check if anchor was in original selection (already counted)
                                var wasInSelection = false;
                                for (var selIdx = 0; selIdx < selection.length; selIdx++) {
                                    if (selection[selIdx] === centerAnchors[anchorIdx]) {
                                        wasInSelection = true;
                                        break;
                                    }
                                }

                                if (!wasInSelection) {
                                    $.writeln("[MOVE]     Moving anchor from '" + anchorLayer + "' to '" + targetLayerName + "'");
                                    centerAnchors[anchorIdx].move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                                    anchorsMoved++;
                                } else {
                                    $.writeln("[MOVE]     Anchor was in selection, already counted");
                                }
                            } catch (eAnchor) {
                                $.writeln("[MOVE]     Error moving anchor: " + eAnchor);
                            }
                        }
                    }
                }
                // Move PathItems (anchors)
                else if (item.typename === 'PathItem') {
                    $.writeln("[MOVE]   Moving PathItem with " + item.pathPoints.length + " points...");
                    item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                    if (item.pathPoints.length === 1) {
                        anchorsMoved++;
                        $.writeln("[MOVE]   Anchor moved");
                    } else {
                        itemsMoved++;
                        $.writeln("[MOVE]   Path moved");
                    }
                }
                // Move GroupItems
                else if (item.typename === 'GroupItem') {
                    $.writeln("[MOVE]   Moving GroupItem...");
                    item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                    itemsMoved++;
                    $.writeln("[MOVE]   GroupItem moved");
                }
                // Move other items
                else {
                    $.writeln("[MOVE]   Moving " + item.typename + "...");
                    item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                    itemsMoved++;
                    $.writeln("[MOVE]   Item moved");
                }
            } catch (eMove) {
                $.writeln("[MOVE]   ERROR moving item: " + eMove);
            }
        }

        $.writeln("[MOVE] Moved " + itemsMoved + " items, " + anchorsMoved + " anchors, skipped " + itemsSkipped + " items");

        // Write debug log to file
        try {
            var debugLogPath = "C:/Users/Chris/AppData/Roaming/Adobe/CEP/extensions/Magic-Ductwork-Panel/move-debug.log";
            var debugFile = new File(debugLogPath);
            if (debugFile.open("w")) {
                debugFile.writeln("Move operation completed:");
                debugFile.writeln("Items moved: " + itemsMoved);
                debugFile.writeln("Anchors moved: " + anchorsMoved);
                debugFile.writeln("Items skipped: " + itemsSkipped);
                debugFile.writeln("Target layer: " + targetLayerName);
                debugFile.close();
            }
        } catch (eDebug) {
            $.writeln("[MOVE] Error writing debug log: " + eDebug);
        }

        // Lock and hide Ignore layer if that's the target
        if (isIgnoreLayer) {
            targetLayer.locked = true;
            targetLayer.visible = false;
            $.writeln("[MOVE] Locked and hid Ignore layer");
        } else {
            // Restore original state for non-Ignore layers
            targetLayer.locked = wasLocked;
            targetLayer.visible = wasVisible;
            $.writeln("[MOVE] Restored layer state");
        }

        doc.selection = null;

        var result = JSON.stringify({
            itemsMoved: itemsMoved,
            anchorsMoved: anchorsMoved,
            itemsSkipped: itemsSkipped
        });
        $.writeln("[MOVE] Returning result: " + result);
        return result;
    } catch (e) {
        $.writeln("[MOVE] EXCEPTION: " + e.toString());
        $.writeln("[MOVE] Stack: " + e.line);
        return "ERROR:" + e.toString();
    }
}
