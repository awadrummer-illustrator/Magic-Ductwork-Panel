#target illustrator

#include "./register-ignore.jsx"

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

function MDUX_runEmoryDuctwork() {
    try {
        if (!app.documents.length) {
            return "ERROR:No document open";
        }

        var doc = app.activeDocument;
        var REGISTER_WIRE_TAG = "MD:REGISTER_WIRE";
        var WIRE_CONNECTION_TOLERANCE = 50; // Increased tolerance to 50px
        var swappedCount = 0;
        var wireCount = 0;

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

        // Step 2: Generate register wires
        var ductworkPaths = [];
        var registerPoints = [];

        $.writeln("[EMORY] Scanning for ductwork paths and registers");

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

        $.writeln("[EMORY] Found " + ductworkPaths.length + " ductwork paths");
        $.writeln("[EMORY] Found " + registerPoints.length + " register points");

        for (var i = 0; i < ductworkPaths.length; i++) {
            var path = ductworkPaths[i];
            if (!path.pathPoints || path.pathPoints.length < 2) continue;
            var layer = path.layer;
            if (!layer) continue;

            var endpoints = [
                { x: path.pathPoints[0].anchor[0], y: path.pathPoints[0].anchor[1] },
                { x: path.pathPoints[path.pathPoints.length - 1].anchor[0], y: path.pathPoints[path.pathPoints.length - 1].anchor[1] }
            ];

            for (var j = 0; j < endpoints.length; j++) {
                var endpoint = endpoints[j];
                var closestRegister = null;
                var closestDist = WIRE_CONNECTION_TOLERANCE;

                for (var k = 0; k < registerPoints.length; k++) {
                    var register = registerPoints[k];
                    var dx = register.x - endpoint.x;
                    var dy = register.y - endpoint.y;
                    var dist = Math.sqrt(dx * dx + dy * dy);
                    if (dist < closestDist) {
                        closestDist = dist;
                        closestRegister = register;
                    }
                }

                if (closestRegister) {
                    try {
                        $.writeln("[EMORY] Creating wire from [" + endpoint.x.toFixed(1) + "," + endpoint.y.toFixed(1) + "] to [" + closestRegister.x.toFixed(1) + "," + closestRegister.y.toFixed(1) + "] (dist=" + closestDist.toFixed(1) + ")");

                        var wirePath = layer.pathItems.add();
                        wirePath.setEntirePath([[endpoint.x, endpoint.y], [closestRegister.x, closestRegister.y]]);
                        wirePath.closed = false;
                        wirePath.stroked = true;
                        wirePath.strokeWidth = 3;
                        var wireColor = new RGBColor();
                        wireColor.red = 0;
                        wireColor.green = 0;
                        wireColor.blue = 255;
                        wirePath.strokeColor = wireColor;
                        wirePath.strokeCap = StrokeCap.ROUNDENDCAP;
                        wirePath.strokeJoin = StrokeJoin.ROUNDENDJOIN;
                        wirePath.filled = false;

                        var wirePoints = wirePath.pathPoints;
                        if (wirePoints && wirePoints.length === 2) {
                            var wireDX = closestRegister.x - endpoint.x;
                            var wireDY = closestRegister.y - endpoint.y;
                            var wireLen = Math.sqrt(wireDX * wireDX + wireDY * wireDY);
                            if (wireLen > 0) {
                                var handleLen = Math.min(wireLen * 0.25, 20);
                                var startPoint = wirePoints[0];
                                startPoint.leftDirection = [endpoint.x, endpoint.y];
                                startPoint.rightDirection = [endpoint.x - handleLen, endpoint.y];
                                startPoint.pointType = PointType.SMOOTH;
                                var endPoint = wirePoints[1];
                                endPoint.leftDirection = [closestRegister.x, closestRegister.y + handleLen * 0.6];
                                endPoint.rightDirection = [closestRegister.x, closestRegister.y];
                                endPoint.pointType = PointType.SMOOTH;
                            }
                        }

                        wirePath.name = "Register Wire";
                        wirePath.note = REGISTER_WIRE_TAG;
                        wireCount++;
                    } catch (eWire) {
                        $.writeln("[EMORY] Wire creation error: " + eWire);
                    }
                }
            }
        }

        $.writeln("[EMORY] Complete: Swapped " + swappedCount + " items, created " + wireCount + " wires");
        return "Swapped " + swappedCount + " items to Emory versions and created " + wireCount + " register wires.";
    } catch (e) {
        $.writeln("[EMORY] Error: " + e);
        return "ERROR:" + e.toString();
    }
}
