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
    } catch (e) { }
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
    try { $.writeln("[MDUX] Updated MDUX_JSX_FOLDER from MDUX_LAST_BRIDGE_PATH: " + $.global.MDUX_JSX_FOLDER); } catch (e) { }
} else if (typeof $.global.MDUX_JSX_FOLDER === "undefined") {
    $.global.MDUX_JSX_FOLDER = File($.fileName).parent;
    try { $.writeln("[MDUX] Initial setup from $.fileName: " + $.global.MDUX_JSX_FOLDER); } catch (e) { }
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
                .replace(/[\b]/g, '\\b') + '"';
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

// Ductwork line layers (line-only layers that should not be orthogonalized, etc.)
var DUCTWORK_LINES = [
    "Green Ductwork",
    "Light Green Ductwork",
    "Blue Ductwork",
    "Orange Ductwork",
    "Light Orange Ductwork",
    "Thermostat Lines"
];

// Ductwork piece layers (used for placing pieces and anchor checks)
var DUCTWORK_PARTS = [
    "Thermostats",
    "Units",
    "Secondary Exhaust Registers",
    "Exhaust Registers",
    "Orange Register",
    "Rectangular Registers",
    "Square Registers",
    "Circular Registers"
];

function MDUX_isDuctworkLine(item) {
    try {
        if (!item || !item.layer) return false;
        var layerName = item.layer.name;
        for (var i = 0; i < DUCTWORK_LINES.length; i++) {
            if (layerName === DUCTWORK_LINES[i]) return true;
        }
    } catch (e) { }
    return false;
}

function MDUX_isDuctworkPart(item) {
    try {
        if (!item || !item.layer) return false;
        var layerName = item.layer.name;
        for (var i = 0; i < DUCTWORK_PARTS.length; i++) {
            if (layerName === DUCTWORK_PARTS[i]) return true;
        }
    } catch (e) { }
    return false;
}

// Metadata storage using item.note (more reliable than tags with undo/redo)
function MDUX_getMetadata(item) {
    try {
        var note = item.note || "";
        if (!note || note.indexOf("MDUX_META:") !== 0) return null;
        var jsonStr = note.substring(10); // Remove "MDUX_META:" prefix
        MDUX_debugLog("[META GET] Raw note length=" + note.length + ", jsonStr length=" + jsonStr.length);
        MDUX_debugLog("[META GET] jsonStr=" + jsonStr.substring(0, 500));

        // Fix corrupted metadata: strip any trailing chars after the closing brace
        var lastBrace = jsonStr.lastIndexOf("}");
        if (lastBrace !== -1 && lastBrace < jsonStr.length - 1) {
            var trailing = jsonStr.substring(lastBrace + 1);
            MDUX_debugLog("[META GET] WARNING: Found trailing chars after JSON: '" + trailing + "' - stripping them");
            jsonStr = jsonStr.substring(0, lastBrace + 1);
        }

        var parsed = JSON.parse(jsonStr);
        // List keys manually (ExtendScript doesn't have Object.keys)
        var keyList = [];
        for (var k in parsed) {
            if (parsed.hasOwnProperty(k)) keyList.push(k);
        }
        MDUX_debugLog("[META GET] Parsed keys: " + keyList.join(", "));
        return parsed;
    } catch (e) {
        MDUX_debugLog("[META GET] ERROR parsing JSON: " + e + " | jsonStr was: " + (jsonStr || "(undefined)").substring(0, 200));
        return null;
    }
}

function MDUX_setMetadata(item, metadata) {
    try {
        var jsonStr = JSON.stringify(metadata);
        var noteStr = "MDUX_META:" + jsonStr;
        item.note = noteStr;

        // Verify it was actually set
        var verify = item.note;
        MDUX_debugLog("[META] Set metadata on " + item.typename + ": " + noteStr.substring(0, 500));
        MDUX_debugLog("[META] Verified note = " + (verify ? verify.substring(0, 500) : "NULL"));

        if (!verify || verify !== noteStr) {
            MDUX_debugLog("[META] ERROR: Note was not set correctly!");
        }
    } catch (e) {
        MDUX_debugLog("[META] ERROR setting metadata: " + e);
    }
}

function MDUX_getTag(item, key) {
    var meta = MDUX_getMetadata(item);
    return meta && meta[key] !== undefined ? String(meta[key]) : null;
}

function MDUX_setTag(item, key, value) {
    var meta = MDUX_getMetadata(item) || {};
    meta[key] = value;
    MDUX_setMetadata(item, meta);
}

function MDUX_removeTag(item, key) {
    var meta = MDUX_getMetadata(item);
    if (!meta) return;
    delete meta[key];
    MDUX_setMetadata(item, meta);
}

// ========================================
// GLOBAL DOCUMENT SCALE MANAGEMENT
// ========================================

/**
 * Gets or creates the Scale Factor Box used to store the document's scale factor.
 * Logic ported from the user's original script.
 */
function MDUX_getOrCreateScaleFactorBox(doc) {
    var layerName = "Scale Factor Container Layer";
    var boxName = "ScaleFactorBox";
    var container;

    try {
        container = doc.layers.getByName(layerName);
    } catch (e) {
        container = doc.layers.add();
        container.name = layerName;
        container.printable = false; // Usually don't want this to print
    }

    var box = null;
    for (var i = 0; i < container.pathItems.length; i++) {
        var pi = container.pathItems[i];
        if (pi.name === boxName) {
            box = pi;
            break;
        }
    }

    if (!box) {
        // Create initial box
        doc.rulerOrigin = [0, 0];
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var left = ab.artboardRect[0];
        var top = ab.artboardRect[1];
        var w = 100, h = 100;

        // Temporarily unlock container if needed
        var wasLocked = container.locked;
        container.locked = false;

        box = container.pathItems.rectangle(top - h, left - 125, w, h);
        box.position = [left - 125, top];
        box.name = boxName;
        box.filled = false;
        box.stroked = true;
        box.strokeWidth = 1;

        var blue = new RGBColor();
        blue.red = 0; blue.green = 0; blue.blue = 255;
        box.strokeColor = blue;
        box.note = "100"; // Default scale

        container.locked = wasLocked;
    }

    return box;
}

function MDUX_getDocumentScale() {
    try {
        if (app.documents.length === 0) return "100";
        var box = MDUX_getOrCreateScaleFactorBox(app.activeDocument);
        return box.note || "100";
    } catch (e) {
        return "100";
    }
}

// LEGACY: Document scale editing disabled in UI to prevent desync issues.
// This function is kept for internal use only (called by MDUX_applyScaleToFullDocument).
function MDUX_setDocumentScale(percent) {
    try {
        if (app.documents.length === 0) return "ERROR:No document";
        var doc = app.activeDocument;
        var box = MDUX_getOrCreateScaleFactorBox(doc);

        var container = box.layer;
        var wasLocked = container.locked;
        container.locked = false;

        box.note = String(percent);

        container.locked = wasLocked;
        return "OK";
    } catch (e) {
        return "ERROR:" + e;
    }
}

/**
 * LEGACY: Full document scaling disabled in UI to prevent desync issues.
 * Kept for reference - do not expose in panel UI.
 *
 * Ported logic from old script: Portions the document into Geometry (Resized)
 * and Strokes (StrokeWidth Scaled).
 */
function MDUX_applyScaleToFullDocument(targetPercent) {
    try {
        if (app.documents.length === 0) return JSON.stringify({ ok: false, message: "No document open" });
        var doc = app.activeDocument;

        var currentScaleStr = MDUX_getDocumentScale();
        var currentPercent = parseFloat(currentScaleStr) || 100;

        if (Math.abs(targetPercent - currentPercent) < 0.001) {
            return JSON.stringify({ ok: true, message: "Document is already at " + targetPercent + "%" });
        }

        var ratio = targetPercent / currentPercent;
        var stats = { parts: 0, lines: 0 };

        // 1. Scale Parts (Geometry & Stroke)
        for (var i = 0; i < DUCTWORK_PARTS.length; i++) {
            try {
                var layer = doc.layers.getByName(DUCTWORK_PARTS[i]);
                if (layer.locked || !layer.visible) continue;

                for (var j = 0; j < layer.pageItems.length; j++) {
                    var item = layer.pageItems[j];
                    if (item.locked || item.hidden) continue;
                    try {
                        // Resizing parts includes stroke scaling usually
                        item.resize(ratio * 100, ratio * 100, true, true, true, true, ratio * 100, Transformation.CENTER);
                        stats.parts++;
                    } catch (e) { }
                }
            } catch (eLayer) { }
        }

        // 2. Scale Lines (Stroke Only)
        // Note: For lines, we typically want to scale the stroke width but not the geometry lengths
        // unless they are also being repositioned. The old script's applyFullLayerScaling
        // distinguishes between scaleGeometryOnly and scaleStrokesOnly.

        for (var k = 0; k < DUCTWORK_LINES.length; k++) {
            try {
                var lLayer = doc.layers.getByName(DUCTWORK_LINES[k]);
                if (lLayer.locked || !lLayer.visible) continue;

                for (var m = 0; m < lLayer.pageItems.length; m++) {
                    var lItem = lLayer.pageItems[m];
                    if (lItem.locked || lItem.hidden) continue;

                    // Recursive stroke scaling logic
                    MDUX_scaleStrokeRecursive(lItem, ratio);
                    stats.lines++;
                }
            } catch (eLayerLine) { }
        }

        // Update the box
        MDUX_setDocumentScale(targetPercent);

        // SYNC TAGS: Important to keep Selection Transform UI in sync
        for (var n = 0; n < doc.pageItems.length; n++) {
            var pi = doc.pageItems[n];
            // Only update if it already has ductwork metadata to avoid tagging every single thing in the doc
            if (MDUX_getTag(pi, "MDUX_OriginalWidth") !== null) {
                MDUX_setTag(pi, "MDUX_CurrentScale", targetPercent);
            }
        }

        return JSON.stringify({ ok: true, message: "Scaled document from " + currentPercent + "% to " + targetPercent + "%", stats: stats });
    } catch (e) {
        return JSON.stringify({ ok: false, message: "Error: " + e });
    }
}

function MDUX_scaleStrokeRecursive(item, ratio) {
    var percent = ratio * 100;
    try {
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                MDUX_scaleStrokeRecursive(item.pageItems[i], ratio);
            }
        } else if (item.typename === "CompoundPathItem" || item.typename === "PathItem") {
            // Use resize(100, 100, false, false, false, false, percent) 
            // to scale ONLY the stroke width/patterns.
            // Ported from old script's scaleStrokeProperties logic.
            item.resize(100, 100, false, false, false, false, percent, Transformation.CENTER);
        }
    } catch (e) { }
}

function MDUX_resetTransforms(targetPercent) {
    try {
        if (app.documents.length === 0) return JSON.stringify({ ok: false, message: "No document open." });
        var sel = app.selection;
        if (!sel || sel.length === 0) return JSON.stringify({ ok: false, message: "Nothing selected." });

        var count = 0;
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];

            // 1. Restore Rotation
            var cumRot = MDUX_getTag(item, "MDUX_CumulativeRotation");
            if (cumRot !== null) {
                var r = parseFloat(cumRot);
                if (!isNaN(r) && r !== 0) {
                    item.rotate(-r, true, true, true, true, Transformation.CENTER);
                }
                MDUX_removeTag(item, "MDUX_CumulativeRotation");
            }

            // 2. Restore Dimensions (Width/Height)
            var origW = MDUX_getTag(item, "MDUX_OriginalWidth");
            var origH = MDUX_getTag(item, "MDUX_OriginalHeight");
            if (origW !== null && origH !== null) {
                item.width = parseFloat(origW);
                item.height = parseFloat(origH);
                MDUX_removeTag(item, "MDUX_OriginalWidth");
                MDUX_removeTag(item, "MDUX_OriginalHeight");
            }

            // 3. Restore Stroke Width
            var origStroke = MDUX_getTag(item, "MDUX_OriginalStrokeWidth");
            if (origStroke !== null) {
                item.strokeWidth = parseFloat(origStroke);
                MDUX_removeTag(item, "MDUX_OriginalStrokeWidth");
            }

            // 4. Restore Selection Transform Tag
            MDUX_removeTag(item, "MDUX_CurrentScale");

            count++;
        }
        return JSON.stringify({ ok: true, message: "Reset " + count + " items." });
    } catch (e) {
        return JSON.stringify({ ok: false, message: "Error: " + e.message });
    }
}

if (typeof MDUX === "undefined") {
    var MDUX = {};
}

// Debug logging to file
// In-memory debug log buffer
if (typeof $.global.MDUX_debugBuffer === "undefined") {
    $.global.MDUX_debugBuffer = [];
}

function MDUX_debugLog(message) {
    try {
        // ExtendScript doesn't have toISOString(), use toString() instead
        var timestamp = new Date().toString();
        var logEntry = "[" + timestamp + "] " + message;
        $.global.MDUX_debugBuffer.push(logEntry);

        // Keep only last 200 entries to prevent memory issues
        if ($.global.MDUX_debugBuffer.length > 200) {
            $.global.MDUX_debugBuffer.shift();
        }
    } catch (e) {
        // Log the error to help debug
        $.global.MDUX_debugBuffer.push("[ERROR in MDUX_debugLog] " + e.toString());
    }
}

// Log that bridge is loading
MDUX_debugLog("=== BRIDGE LOADING - panel-bridge.jsx ===");

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
        try { $.writeln("[MDUX] root folder: " + rootPath); } catch (logRoot) { }

        var scriptFile = File(MDUX_joinPath(root, "magic-final.jsx"));
        var scriptPath = "";
        try {
            scriptPath = scriptFile.fsName || scriptFile.toString();
        } catch (e) {
            scriptPath = String(scriptFile);
        }
        MDUX_debugLog("looking for: " + scriptPath);
        MDUX_debugLog("file exists: " + scriptFile.exists);
        try { $.writeln("[MDUX] looking for: " + scriptPath); } catch (logPath) { }
        try { $.writeln("[MDUX] file exists: " + scriptFile.exists); } catch (logExists) { }

        if (!scriptFile.exists) {
            MDUX_debugLog("ERROR: magic-final missing at " + scriptPath);
            try { $.writeln("[MDUX] magic-final missing at " + scriptPath); } catch (logMiss) { }
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
            try { $.writeln("[MDUX] requireMagicFinal: could not determine root"); } catch (logErr) { }
            return false;
        }
        var scriptFile = File(MDUX_joinPath(root, "magic-final.jsx"));
        var scriptPath = "";
        try {
            scriptPath = scriptFile.fsName || scriptFile.toString();
        } catch (e) {
            scriptPath = String(scriptFile);
        }
        try { $.writeln("[MDUX] requireMagicFinal: looking for " + scriptPath); } catch (logPath) { }
        if (!scriptFile.exists) {
            try { $.writeln("[MDUX] requireMagicFinal: file not found at " + scriptPath); } catch (logMiss) { }
            return false;
        }
        var previousForced = null;
        if ($.global.MDUX && $.global.MDUX.forcedOptions) {
            previousForced = $.global.MDUX.forcedOptions;
        }
        $.global.MDUX = $.global.MDUX || {};
        $.global.MDUX.forcedOptions = { action: "library" };
        try { $.writeln("[MDUX] requireMagicFinal: loading magic-final.jsx"); } catch (logErr) { }
        $.evalFile(scriptFile);
        try {
            if ($.global.MDUX) {
                MDUX = $.global.MDUX;
            }
        } catch (eAssign) { }
        if (previousForced) {
            $.global.MDUX.forcedOptions = previousForced;
        } else if ($.global.MDUX && $.global.MDUX.hasOwnProperty("forcedOptions")) {
            delete $.global.MDUX.forcedOptions;
        }
        var ns = (typeof MDUX !== "undefined" && MDUX.rotateSelection) ? MDUX : ($.global.MDUX || null);
        return !!(ns && (ns.rotateSelection || ns.createStandardLayerBlock || ns.importDuctworkGraphicStyles));
    } catch (e) {
        try { $.writeln("[MDUX] requireMagicFinal error: " + e); } catch (logErr2) { }
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
            } else if (item.typename === "CompoundPathItem") {
                // Clear rotation override from child paths (MD:ROT= tokens)
                if (item.pathItems) {
                    for (var cpi = 0; cpi < item.pathItems.length; cpi++) {
                        MDUX.clearRotationOverride(item.pathItems[cpi]);
                        count++;
                    }
                }
                // Also clear rotation override from compound path's MDUX_META
                try {
                    var compoundMeta = MDUX_getMetadata(item);
                    if (compoundMeta && compoundMeta.MDUX_RotationOverride !== undefined) {
                        delete compoundMeta.MDUX_RotationOverride;
                        MDUX_setMetadata(item, compoundMeta);
                    }
                } catch (eCompoundMeta) {
                    // Ignore metadata errors
                }
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

function MDUX_scaleSelectionBridge(requestedPercent) {
    try {
        var percent = requestedPercent;
        if (typeof percent !== "number") percent = parseFloat(percent);
        if (!isFinite(percent)) return "ERROR:Invalid scale value";
        if (!MDUX_requireMagicFinal()) {
            return "ERROR:Scale function unavailable";
        }
        if (typeof MDUX !== "undefined" && MDUX.scaleSelectionAbsolute) {
            // Anchor-based scaling: scale relative to document's anchor
            var anchor = parseFloat(MDUX_getDocumentScale()) || 100;
            var targetPercent = percent * (anchor / 100);

            var stats = MDUX.scaleSelectionAbsolute(targetPercent);

            // Note: We do NOT update the Document Scale anchor here, 
            // as this is a selection-only transform.

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

            // Sync with Scale Factor Box if successful
            if (stats && stats.reset > 0) {
                MDUX_setDocumentScale(100);
            }

            return JSON.stringify(stats);
        }
        return "ERROR:Reset scale function unavailable";
    } catch (e) {
        return "ERROR:" + e;
    }
}

function MDUX_rotationStateBridge() {
    try {
        MDUX_debugLog("[ROT-BRIDGE] MDUX_rotationStateBridge called");
        if (!MDUX_requireMagicFinal()) {
            MDUX_debugLog("[ROT-BRIDGE] MDUX_requireMagicFinal returned false");
            return "ERROR:Rotation function unavailable";
        }
        MDUX_debugLog("[ROT-BRIDGE] MDUX_requireMagicFinal returned true, checking MDUX.getRotationOverrideSummary...");
        if (typeof MDUX !== "undefined" && MDUX.getRotationOverrideSummary) {
            MDUX_debugLog("[ROT-BRIDGE] Calling MDUX.getRotationOverrideSummary()...");
            var summary = MDUX.getRotationOverrideSummary();
            MDUX_debugLog("[ROT-BRIDGE] Got summary: " + JSON.stringify(summary).substring(0, 200));
            return JSON.stringify(summary);
        }
        MDUX_debugLog("[ROT-BRIDGE] MDUX.getRotationOverrideSummary not available");
        return "ERROR:Rotation function unavailable";
    } catch (e) {
        MDUX_debugLog("[ROT-BRIDGE] ERROR: " + e);
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
        try { $.writeln("[MDUX] Import styles: opening " + sourceFile.fsName); } catch (logOpen) { }
        try {
            sourceDoc = app.open(sourceFile);
            app.activeDocument = sourceDoc;
            try {
                for (var L = 0; L < sourceDoc.layers.length; L++) {
                    try { sourceDoc.layers[L].locked = false; } catch (eLock) { }
                    try { sourceDoc.layers[L].visible = true; } catch (eVis) { }
                }
            } catch (eIter) { }

            var items = null;
            try { items = sourceDoc.pageItems; } catch (eItems) { items = null; }
            if (!items || items.length === 0) {
                sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
                app.activeDocument = destDoc;
                return "ERROR:Source document contained no artwork.";
            }

            for (var i = 0; i < items.length; i++) {
                try { items[i].selected = true; } catch (eSel) { }
            }
            app.copy();
            sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (copyErr) {
            try {
                if (sourceDoc) sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeErr) { }
            app.activeDocument = destDoc;
            try { $.writeln("[MDUX] Import styles error: " + copyErr); } catch (logCopy) { }
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
        try { tempLayer.locked = false; } catch (lockErr) { }
        destDoc.activeLayer = tempLayer;
        app.executeMenuCommand("pasteInPlace");
        var pasted = null;
        try { pasted = destDoc.selection; } catch (selErr) { pasted = null; }
        if (pasted) {
            if (pasted.length === undefined) pasted = [pasted];
            for (var p = pasted.length - 1; p >= 0; p--) {
                try { pasted[p].remove(); } catch (remErr) { }
            }
        }
        destDoc.selection = null;
        try {
            if (!tempLayer.pageItems.length && !tempLayer.groupItems.length) {
                tempLayer.remove();
            }
        } catch (cleanupErr) { }
        try { app.redraw(); } catch (redErr) { }
        try { $.writeln("[MDUX] Import styles completed successfully"); } catch (logDone) { }
        return "Graphic styles imported.";
    } catch (e) {
        try { $.writeln("[MDUX] Import styles exception: " + e); } catch (logFinal) { }
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
                    try { $.writeln("[MDUX] Created missing layer: " + name); } catch (logCreate) { }
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
                try { prevLocked = target.locked; target.locked = false; } catch (eLock) { }
                try { target.visible = true; } catch (eVisible) { }
                try {
                    target.move(doc.layers[0], ElementPlacement.PLACEBEFORE);
                } catch (eMove) { }
                try {
                    if (prevLocked !== null) target.locked = prevLocked;
                } catch (eRestore) { }
            } catch (eTarget) { }
        }

        try { app.redraw(); } catch (eRedraw) { }
        try { $.writeln("[MDUX] Ensure standard layers completed"); } catch (logDone) { }
        return "Standard ductwork layers ensured.";
    } catch (e) {
        try { $.writeln("[MDUX] Create layers exception: " + e); } catch (logErr) { }
        return "ERROR:Create layers (exception): " + e;
    }
}

/**
 * Helper function to set the proper name for a PlacedItem based on its file.
 * Extracts the piece name from the file path and sets it as "PieceName (Linked)"
 * This ensures the item can be found by existingItems filters in magic-final.jsx
 */
function MDUX_setPlacedItemName(item, file) {
    try {
        if (!item || !file) return;

        // Get file name without extension
        var fileName = file.name || "";
        var baseName = fileName.replace(/\.[^.]*$/, ""); // Remove extension

        // Remove " Emory" suffix if present (so "Unit Emory" becomes "Unit")
        baseName = baseName.replace(/ Emory$/, "");

        // Set name in format "PieceName (Linked)" to match magic-final.jsx convention
        item.name = baseName + " (Linked)";

        $.writeln("[NAME] Set PlacedItem name to: " + item.name);
    } catch (e) {
        $.writeln("[NAME] Error setting PlacedItem name: " + e);
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
                                MDUX_setPlacedItemName(item, emoryFile);
                                swappedCount++;
                            }
                        }
                    }
                } catch (e) { }
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
            } catch (e) { }

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
                    } catch (e) { }
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
                                MDUX_setPlacedItemName(foundPlacedItem, emoryRegisterFile);
                                swappedCount++;
                                $.writeln("[EMORY] Swapped register at [" + anchorPos.x.toFixed(1) + "," + anchorPos.y.toFixed(1) + "] to " + targetImageName);
                            }
                        } else {
                            // Place new Emory register at anchor
                            var newPlaced = registerLayer.placedItems.add();
                            newPlaced.file = emoryRegisterFile;
                            MDUX_setPlacedItemName(newPlaced, emoryRegisterFile);
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
                    try { layerName = item.layer.name; } catch (e) { }
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
                } catch (e) { }

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
            } catch (e) { }

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
                            } catch (e) { }

                            wirePath.stroked = true;
                            wirePath.strokeWidth = 3;
                            try {
                                var wireColor = new RGBColor();
                                wireColor.red = 0;
                                wireColor.green = 0;
                                wireColor.blue = 255;
                                wirePath.strokeColor = wireColor;
                            } catch (eWireColor) { }
                            wirePath.strokeCap = StrokeCap.ROUNDENDCAP;
                            wirePath.strokeJoin = StrokeJoin.ROUNDENDJOIN;

                            wirePath.filled = false;
                            try {
                                var noColor = new NoColor();
                                wirePath.fillColor = noColor;
                            } catch (e) { }

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
                                    try { startPointInfo.pointType = PointType.SMOOTH; } catch (eWireStartType) { }
                                }

                                // End handle: point downward (inverted)
                                if (endPointInfo) {
                                    endPointInfo.leftDirection = [
                                        endPoint.x,
                                        endPoint.y + handleLen * 0.6
                                    ];
                                    endPointInfo.rightDirection = [endPoint.x, endPoint.y];
                                    try { endPointInfo.pointType = PointType.SMOOTH; } catch (eWireEndType) { }
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
        var artPlacedPositions = []; // Track positions where art has been placed to avoid duplicates

        // Helper to check if art was already placed at a position (within 5px tolerance)
        function wasArtPlacedNear(x, y) {
            for (var ap = 0; ap < artPlacedPositions.length; ap++) {
                var dx = artPlacedPositions[ap].x - x;
                var dy = artPlacedPositions[ap].y - y;
                if (Math.sqrt(dx * dx + dy * dy) < 5) {
                    return true;
                }
            }
            return false;
        }

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

            // Check if item is on a valid source layer (ductwork parts or Ignore only - NOT ductwork lines)
            var itemLayerName = item.layer ? item.layer.name : null;
            $.writeln("[MOVE]   Item layer: " + itemLayerName);

            var isOnDuctworkParts = isValidDuctworkLayer(itemLayerName);
            var isOnIgnoreLayer = (itemLayerName === 'Ignore' || itemLayerName === 'Ignored');
            var isOnValidSourceLayer = isOnDuctworkParts || isOnIgnoreLayer;

            if (!isOnValidSourceLayer) {
                $.writeln("[MOVE]   SKIPPED - not on a ductwork parts or Ignore layer (ductwork lines excluded)");
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
                            MDUX_setPlacedItemName(newItem, filePath);

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

                            // Record this position to prevent duplicate art placement
                            artPlacedPositions.push({ x: targetX, y: targetY });

                            // Write complete metadata for the placed item
                            try {
                                // Get actual dimensions after placement
                                var actualWidth = newItem.width;
                                var actualHeight = newItem.height;
                                var actualStrokeWidth = 1;
                                try { actualStrokeWidth = newItem.strokeWidth || 1; } catch (e) { }

                                // Check if there's a rotation override from the nearest anchor
                                var rotationOverride = null;
                                if (nearestAnchor && nearestAnchor.anchor) {
                                    try {
                                        var anchorMeta = MDUX_getMetadata(nearestAnchor.anchor);
                                        if (anchorMeta && anchorMeta.MDUX_PointRotation !== undefined && anchorMeta.MDUX_PointRotation !== null) {
                                            rotationOverride = parseFloat(anchorMeta.MDUX_PointRotation);
                                            if (!isFinite(rotationOverride)) rotationOverride = null;
                                        }
                                    } catch (eAnchorMeta) {
                                        rotationOverride = null;
                                    }
                                }

                                var metadata = {
                                    MDUX_OriginalWidth: actualWidth,
                                    MDUX_OriginalHeight: actualHeight,
                                    MDUX_OriginalStrokeWidth: actualStrokeWidth,
                                    MDUX_CumulativeRotation: "0",
                                    MDUX_CurrentScale: String(smallestScale),
                                    tagScale: smallestScale,
                                    tagRotation: 0
                                };

                                // Store rotation override if available
                                if (rotationOverride !== null) {
                                    metadata.MDUX_RotationOverride = rotationOverride;
                                }

                                MDUX_setMetadata(newItem, metadata);
                                $.writeln("[MOVE]   Wrote complete metadata: scale=" + smallestScale + ", rotation=0" + (rotationOverride !== null ? ", rotOverride=" + rotationOverride : "") + ", width=" + actualWidth);
                            } catch (eMetadata) {
                                $.writeln("[MOVE]   Warning: Failed to write metadata: " + eMetadata);
                            }

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
                            } catch (e) { }
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
                // Move PathItems (anchors or multi-point paths)
                else if (item.typename === 'PathItem') {
                    var numPoints = item.pathPoints.length;
                    $.writeln("[MOVE]   Processing PathItem with " + numPoints + " points...");

                    // Handle paths with art placement (not Ignore layer, and we have a file)
                    if (filePath && !isIgnoreLayer) {

                        // Collect all anchor positions from the path
                        var anchorPositions = [];
                        for (var pi = 0; pi < numPoints; pi++) {
                            var pos = item.pathPoints[pi].anchor;
                            anchorPositions.push({ x: pos[0], y: pos[1] });
                        }
                        $.writeln("[MOVE]   Extracted " + anchorPositions.length + " anchor positions");

                        // Delete the original path (we'll create individual anchors)
                        item.remove();
                        $.writeln("[MOVE]   Original path removed");

                        // For each anchor position, create an anchor point and place art (if not already placed)
                        for (var ai = 0; ai < anchorPositions.length; ai++) {
                            var anchorX = anchorPositions[ai].x;
                            var anchorY = anchorPositions[ai].y;
                            $.writeln("[MOVE]   Processing anchor " + (ai + 1) + "/" + anchorPositions.length + " at " + anchorX.toFixed(2) + ", " + anchorY.toFixed(2));

                            // Create a new single-point anchor on the target layer
                            var newAnchor = targetLayer.pathItems.add();
                            newAnchor.setEntirePath([[anchorX, anchorY]]);
                            newAnchor.filled = false;
                            newAnchor.stroked = false;
                            anchorsMoved++;

                            // Check if art was already placed at this position (avoid duplicates)
                            if (wasArtPlacedNear(anchorX, anchorY)) {
                                $.writeln("[MOVE]   SKIPPED art placement - art already placed nearby");
                                continue;
                            }

                            // Place art at this anchor
                            var newItem = targetLayer.placedItems.add();
                            newItem.file = filePath;
                            MDUX_setPlacedItemName(newItem, filePath);

                            // Get new item bounds
                            var newBounds = newItem.geometricBounds;
                            var newWidth = Math.abs(newBounds[2] - newBounds[0]);
                            var newHeight = Math.abs(newBounds[1] - newBounds[3]);

                            // Scale if needed
                            if (smallestScale !== 100) {
                                newItem.resize(smallestScale, smallestScale, true, false, false, false, 100, Transformation.CENTER);
                                newBounds = newItem.geometricBounds;
                                newWidth = Math.abs(newBounds[2] - newBounds[0]);
                                newHeight = Math.abs(newBounds[1] - newBounds[3]);
                            }

                            // Position centered on anchor
                            newItem.position = [anchorX - newWidth / 2, anchorY + newHeight / 2];

                            // Record this position as having art placed
                            artPlacedPositions.push({ x: anchorX, y: anchorY });

                            // Write metadata
                            try {
                                var metadata = {
                                    MDUX_OriginalWidth: newItem.width,
                                    MDUX_OriginalHeight: newItem.height,
                                    MDUX_OriginalStrokeWidth: 1,
                                    MDUX_CumulativeRotation: "0",
                                    MDUX_CurrentScale: String(smallestScale),
                                    tagScale: smallestScale,
                                    tagRotation: 0
                                };
                                MDUX_setMetadata(newItem, metadata);
                            } catch (eMetadata) {}

                            itemsMoved++;
                            $.writeln("[MOVE]   Art placed at anchor " + (ai + 1));
                        }

                    } else {
                        // Just move the path item (no art placement - Ignore layer or no file)
                        item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                        if (numPoints === 1) {
                            anchorsMoved++;
                            $.writeln("[MOVE]   Anchor moved (no art - Ignore layer or no file)");
                        } else {
                            itemsMoved++;
                            $.writeln("[MOVE]   Path moved (no art - Ignore layer or no file)");
                        }
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

function MDUX_scaleStrokeRecursively(item, scalePercent, stats) {
    try {
        if (item.locked || item.hidden) {
            stats.locked++;
            return;
        }

        if (item.typename === "GroupItem") {
            stats.groups++;
            var children = item.pageItems;
            for (var i = 0; i < children.length; i++) {
                MDUX_scaleStrokeRecursively(children[i], scalePercent, stats);
            }
        } else if (item.typename === "CompoundPathItem") {
            stats.compoundPaths++;

            // Try to scale the compound path itself first (if it has stroke properties)
            var compoundStroked = false;
            try {
                // Some AI versions expose stroke on the CompoundPathItem
                if (item.stroked) {
                    item.strokeWidth = item.strokeWidth * (scalePercent / 100.0);
                    stats.stroked++;
                    compoundStroked = true;
                }
            } catch (e) { }

            // Also recurse into children (pathItems)
            // If the compound path itself was stroked, the children might inherit or duplicate.
            // But usually, if the compound path is stroked, the children's stroke properties are ignored or synced.
            // If we didn't successfully scale the compound path, we MUST scale the children.
            // If we DID scale the compound path, scaling children might be redundant but harmless (unless they have independent strokes).

            var paths = item.pathItems;
            for (var j = 0; j < paths.length; j++) {
                MDUX_scaleStrokeRecursively(paths[j], scalePercent, stats);
            }

        } else if (item.typename === "PathItem") {
            stats.paths++;
            var strokesScaled = 0;

            // Scale all strokes from Appearance panel (multiple strokes)
            try {
                if (item.strokes && item.strokes.length > 0) {
                    for (var s = 0; s < item.strokes.length; s++) {
                        var stroke = item.strokes[s];
                        if (stroke.visible) {
                            stroke.weight = stroke.weight * (scalePercent / 100.0);
                            strokesScaled++;
                        }
                    }
                }
            } catch (e) {
                // Fallback: strokes collection not available, use basic strokeWidth
            }

            // Also scale basic strokeWidth if stroked (covers single-stroke case)
            if (item.stroked && strokesScaled === 0) {
                item.strokeWidth = item.strokeWidth * (scalePercent / 100.0);
                strokesScaled++;
            }

            if (strokesScaled > 0) {
                stats.stroked++;
            } else if (item.filled) {
                stats.filled++;
            } else {
                stats.unstroked++;
            }
        } else {
            stats.others++;
            stats.otherTypes.push(item.typename);
        }
    } catch (e) {
        stats.errors++;
    }
}

function MDUX_transformEach(scale, rotation, undoPrevious) {
    MDUX_debugLog("========================================");
    MDUX_debugLog("[TRANSFORM] MDUX_transformEach called: scale=" + scale + ", rotation=" + rotation + ", undoPrevious=" + undoPrevious);
    MDUX_debugLog("========================================");

    try {
        if (app.documents.length === 0) return JSON.stringify({ ok: false, message: "No document open." });

        var sel = app.selection;
        if (!sel || sel.length === 0) return JSON.stringify({ ok: false, message: "Nothing selected." });

        MDUX_debugLog("[TRANSFORM] Selection has " + sel.length + " items");

        var s = Number(scale);
        var r = Number(rotation);

        // Handle Undo AFTER getting selection but BEFORE reading metadata
        // This ensures metadata is preserved across undo/redo cycles
        if (undoPrevious === true || undoPrevious === "true") {
            MDUX_debugLog("[TRANSFORM] Executing undo...");

            // Before undo, check if metadata exists
            var item0 = sel[0];
            var beforeUndo = MDUX_getTag(item0, "MDUX_CurrentScale");
            MDUX_debugLog("[TRANSFORM] Before undo, CurrentScale = " + beforeUndo);

            app.executeMenuCommand('undo');

            // After undo, re-fetch selection in case it changed
            sel = app.selection;
            if (!sel || sel.length === 0) return JSON.stringify({ ok: false, message: "Selection lost after undo." });

            // Check if metadata survived
            item0 = sel[0];
            var afterUndo = MDUX_getTag(item0, "MDUX_CurrentScale");
            MDUX_debugLog("[TRANSFORM] After undo, CurrentScale = " + afterUndo);

            if (beforeUndo && !afterUndo) {
                MDUX_debugLog("[TRANSFORM] WARNING: Metadata was lost during undo!");
            }
        }

        var len = sel.length;
        var transformedCount = 0;
        var stats = {
            groups: 0,
            compoundPaths: 0,
            paths: 0,
            stroked: 0,
            filled: 0,
            unstroked: 0,
            locked: 0,
            errors: 0,
            others: 0,
            otherTypes: []
        };

        var skippedInfo = [];
        for (var i = 0; i < len; i++) {
            var item = sel[i];
            try {
                // Determine item type for scaling logic
                var isDuctLine = MDUX_isDuctworkLine(item);
                var isDuctPart = MDUX_isDuctworkPart(item);

                // Initialize metadata ONLY ONCE when first touched
                // After undo, metadata should still exist from before the transform
                var needsInit = MDUX_getTag(item, "MDUX_OriginalWidth") === null;
                if (needsInit) {
                    try {
                        MDUX_debugLog("[TRANSFORM] Item " + i + " typename=" + item.typename + " needs init.");
                        MDUX_debugLog("[TRANSFORM] Note BEFORE any init: " + (item.note || "(empty)").substring(0, 500));

                        // Safely handle strokeWidth (missing on GroupItems)
                        var sWidth = 1;
                        try { sWidth = item.strokeWidth || 1; } catch (e) { }

                        // Store the ABSOLUTE original dimensions before any transform
                        MDUX_setTag(item, "MDUX_OriginalWidth", item.width);
                        MDUX_debugLog("[TRANSFORM] After OriginalWidth, note: " + (item.note || "(empty)").substring(0, 500));

                        MDUX_setTag(item, "MDUX_OriginalHeight", item.height);
                        MDUX_setTag(item, "MDUX_OriginalStrokeWidth", sWidth);
                        MDUX_setTag(item, "MDUX_CumulativeRotation", "0");
                        MDUX_setTag(item, "MDUX_CurrentScale", "100");

                        MDUX_debugLog("[TRANSFORM] After ALL init tags, note: " + (item.note || "(empty)").substring(0, 500));

                        // Verify tags were actually saved
                        var verifyWidth = MDUX_getTag(item, "MDUX_OriginalWidth");
                        var verifyScale = MDUX_getTag(item, "MDUX_CurrentScale");
                        var verifyRot = MDUX_getTag(item, "MDUX_RotationOverride");
                        MDUX_debugLog("[TRANSFORM] Init complete. Verified: width=" + verifyWidth + ", scale=" + verifyScale + ", rotOverride=" + verifyRot);
                    } catch (eInit) {
                        MDUX_debugLog("[TRANSFORM] ERROR initializing tags: " + eInit +
                                     " for item type=" + item.typename);
                    }
                } else {
                    // Item already has metadata - log it to verify rotation override is preserved
                    MDUX_debugLog("[TRANSFORM] Item " + i + " already has metadata, note: " + (item.note || "(empty)").substring(0, 500));
                    var existingRotOverride = MDUX_getTag(item, "MDUX_RotationOverride");
                    MDUX_debugLog("[TRANSFORM] Existing MDUX_RotationOverride: " + (existingRotOverride || "null/undefined"));
                }

                // Get current state from metadata (or defaults if just initialized)
                var currentScale = parseFloat(MDUX_getTag(item, "MDUX_CurrentScale") || "100");
                var currentRotation = parseFloat(MDUX_getTag(item, "MDUX_CumulativeRotation") || "0");
                var anchor = Number(MDUX_getDocumentScale()) || 100;
                var targetPercent = s * (anchor / 100);

                // --- ROTATION ---
                // r is the DELTA rotation, not absolute
                // We apply the delta and update cumulative
                if (r !== 0 && isDuctPart) {
                    item.rotate(r, true, true, true, true, Transformation.CENTER);
                    MDUX_setTag(item, "MDUX_CumulativeRotation", currentRotation + r);
                }

                // --- SCALING ---
                // s is the target scale on the UI slider (e.g., 110 for 110%)
                // We need to transform from currentScale to targetPercent
                MDUX_debugLog("[TRANSFORM] Item " + i + ": currentScale=" + currentScale +
                         ", targetPercent=" + targetPercent + ", anchor=" + anchor);

                if (Math.abs(currentScale - targetPercent) > 0.001) {
                    var resizeFactor = (targetPercent / currentScale) * 100;
                    MDUX_debugLog("[TRANSFORM] Applying resize factor: " + resizeFactor + "% (isDuctLine=" + isDuctLine + ")");

                    if (isDuctLine) {
                        // Lines: scale stroke only
                        item.resize(100, 100, true, true, true, true, resizeFactor, Transformation.CENTER);
                    } else {
                        // Parts: scale geometry and stroke
                        item.resize(resizeFactor, resizeFactor, true, true, true, true, resizeFactor, Transformation.CENTER);
                    }

                    MDUX_debugLog("[TRANSFORM] Resize complete, setting CurrentScale to " + targetPercent);
                    MDUX_setTag(item, "MDUX_CurrentScale", targetPercent);

                    // Verify it was set
                    var verifySet = MDUX_getTag(item, "MDUX_CurrentScale");
                    MDUX_debugLog("[TRANSFORM] CurrentScale after set: " + verifySet);
                } else {
                    MDUX_debugLog("[TRANSFORM] Scale unchanged (already at target)");
                }
                transformedCount++;
            } catch (eItem) {
                stats.errors++;
            }
        }

        var msg = "Transformed " + transformedCount + "/" + len;
        if (skippedInfo.length > 0) msg += " [Skipped: " + skippedInfo.join(", ") + "]";

        return JSON.stringify({ ok: true, message: msg });

    } catch (e) {
        return JSON.stringify({ ok: false, message: "Error: " + e.message });
    }
}

function MDUX_getSelectedLineAngleBridge() {
    try {
        if (app.documents.length === 0) {
            return JSON.stringify({ ok: false, message: "No document open." });
        }

        var sel = app.selection;
        if (!sel || sel.length === 0) {
            return JSON.stringify({ ok: false, message: "Please select a line." });
        }

        // Find the first PathItem in the selection
        var pathItem = null;
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "PathItem") {
                pathItem = sel[i];
                break;
            }
        }

        if (!pathItem) {
            return JSON.stringify({ ok: false, message: "Please select a path or line." });
        }

        if (!pathItem.pathPoints || pathItem.pathPoints.length < 2) {
            return JSON.stringify({ ok: false, message: "Selected path must have at least 2 points." });
        }

        // Get the first two points of the path
        var point1 = pathItem.pathPoints[0].anchor;
        var point2 = pathItem.pathPoints[1].anchor;

        // Calculate the angle in degrees
        // atan2 returns angle in radians from -PI to PI
        // Negate dy to account for Illustrator's inverted Y-axis (Y increases downward)
        var dx = point2[0] - point1[0];
        var dy = point2[1] - point1[1];
        var angleRadians = Math.atan2(-dy, dx);
        var angleDegrees = angleRadians * (180 / Math.PI);

        // Round to 1 decimal place
        angleDegrees = Math.round(angleDegrees * 10) / 10;

        return JSON.stringify({
            ok: true,
            angle: angleDegrees,
            message: "Angle: " + angleDegrees + "°"
        });

    } catch (e) {
        return JSON.stringify({ ok: false, message: "Error: " + e.message });
    }
}
// Test function to verify note property works
function MDUX_testNoteProperty() {
    try {
        if (app.documents.length === 0) return "ERROR: No document";
        var sel = app.selection;
        if (!sel || sel.length === 0) return "ERROR: Nothing selected";

        var item = sel[0];
        MDUX_debugLog("[TEST] Testing note property on " + item.typename);

        // Test 1: Set a simple note
        item.note = "TEST123";
        var readBack = item.note;
        MDUX_debugLog("[TEST] Set 'TEST123', read back: '" + readBack + "'");

        if (readBack !== "TEST123") {
            return "ERROR: Note property doesn't work on " + item.typename;
        }

        // Test 2: Set metadata
        MDUX_setMetadata(item, { testKey: "testValue", number: 42 });

        // Test 3: Read it back
        var meta = MDUX_getMetadata(item);
        if (!meta) {
            return "ERROR: Metadata not readable";
        }

        MDUX_debugLog("[TEST] Metadata read back: " + JSON.stringify(meta));
        return "SUCCESS: Note property works! Meta=" + JSON.stringify(meta);
    } catch (e) {
        return "ERROR: " + e;
    }
}

// Function to get debug log contents from memory buffer
function MDUX_getDebugLog() {
    try {
        // Add a test message to prove the buffer works
        MDUX_debugLog("MDUX_getDebugLog() was called at " + new Date().toString());

        var bufferInfo = "Buffer type: " + (typeof $.global.MDUX_debugBuffer) +
                        ", Is array: " + ($.global.MDUX_debugBuffer instanceof Array) +
                        ", Length: " + ($.global.MDUX_debugBuffer ? $.global.MDUX_debugBuffer.length : "N/A");

        if (typeof $.global.MDUX_debugBuffer === "undefined") {
            return "DEBUG BUFFER IS UNDEFINED!\n" + bufferInfo;
        }

        if ($.global.MDUX_debugBuffer.length === 0) {
            return "DEBUG BUFFER IS EMPTY (but exists)\n" + bufferInfo + "\nThis means no debug messages have been logged yet.";
        }

        return "=== DEBUG LOG ===\n" + $.global.MDUX_debugBuffer.join("\n");
    } catch (e) {
        return "ERROR reading debug buffer: " + e;
    }
}

// Function to clear debug log buffer
function MDUX_clearDebugLog() {
    try {
        $.global.MDUX_debugBuffer = [];
        return "Debug buffer cleared";
    } catch (e) {
        return "ERROR clearing debug buffer: " + e;
    }
}

function MDUX_getSelectionTransformState() {
    MDUX_debugLog("[SELECT-STATE] Function called");
    try {
        if (app.documents.length === 0) {
            MDUX_debugLog("[SELECT-STATE] No documents open");
            return JSON.stringify({ ok: false });
        }

        var sel = app.selection;
        if (!sel || sel.length === 0) {
            MDUX_debugLog("[SELECT-STATE] No selection");
            return JSON.stringify({ ok: false });
        }
        MDUX_debugLog("[SELECT-STATE] Selection count: " + sel.length);

        var anchor = parseFloat(MDUX_getDocumentScale()) || 100;
        MDUX_debugLog("[SELECT-STATE] Anchor (doc scale): " + anchor);
        var totalScale = 0;
        var totalRot = 0;
        var count = 0;

        var firstScale = null;
        var firstRot = null;
        var mixedScale = false;
        var mixedRot = false;

        for (var i = 0; i < sel.length; i++) {
            try {
                var item = sel[i];

                // Skip anchor points (single-point PathItems) - they shouldn't affect rotation calculations
                if (item.typename === "PathItem") {
                    try {
                        if (item.pathPoints && item.pathPoints.length === 1) {
                            MDUX_debugLog("[SELECT-STATE] Item " + i + ": Skipping anchor (single-point PathItem)");
                            continue;
                        }
                    } catch (eAnchor) {}
                }

                var tagScale = MDUX_getTag(item, "MDUX_CurrentScale");
                var tagRot = MDUX_getTag(item, "MDUX_CumulativeRotation");

                // Fallback: check MDUX_META.MDUX_RotationOverride for placed items
                if ((tagRot === null || tagRot === undefined) && item.typename === "PlacedItem") {
                    try {
                        var meta = MDUX_getMetadata(item);
                        if (meta && meta.MDUX_RotationOverride !== undefined && meta.MDUX_RotationOverride !== null) {
                            tagRot = meta.MDUX_RotationOverride;
                            MDUX_debugLog("[SELECT-STATE] Item " + i + ": Found MDUX_RotationOverride in metadata: " + tagRot);
                        }
                    } catch (eMeta) {}
                }

                MDUX_debugLog("[SELECT-STATE] Item " + i + " (" + item.typename + "): tagScale=" + tagScale + ", tagRot=" + tagRot);

                var absScale = tagScale !== null ? parseFloat(tagScale) : anchor;
                var rot = parseFloat(tagRot || "0");
                MDUX_debugLog("[SELECT-STATE] Item " + i + ": absScale=" + absScale + ", rot=" + rot);

                // Normalize scale relative to anchor for UI
                var uiScale = (absScale / (anchor / 100));
                uiScale = Math.round(uiScale * 100) / 100;
                rot = Math.round(rot * 100) / 100;
                MDUX_debugLog("[SELECT-STATE] Item " + i + ": uiScale=" + uiScale + " (after normalization)");

                if (firstScale === null) {
                    firstScale = uiScale;
                    firstRot = rot;
                } else {
                    if (Math.abs(uiScale - firstScale) > 0.1) mixedScale = true;
                    if (Math.abs(rot - firstRot) > 0.1) mixedRot = true;
                }
                count++;
            } catch (e) {
                MDUX_debugLog("[SELECT-STATE] Error processing item " + i + ": " + e);
            }
        }

        var result = {
            ok: true,
            scale: mixedScale ? null : firstScale,
            rotation: mixedRot ? null : firstRot,
            mixedScale: mixedScale,
            mixedRotation: mixedRot,
            count: count
        };
        MDUX_debugLog("[SELECT-STATE] Returning: " + JSON.stringify(result));

        return JSON.stringify(result);
    } catch (e) {
        MDUX_debugLog("[SELECT-STATE] ERROR: " + e);
        return JSON.stringify({ ok: false, message: e.message });
    }
}

// --- RESET STROKES: Apply appropriate graphic styles to selected ductwork lines based on layer ---
function MDUX_resetStrokes() {
    // Outer safety wrapper
    try {
        if (typeof MDUX_debugLog !== "function") {
            return "Error: MDUX_debugLog not defined";
        }
        MDUX_debugLog("[RESET-STROKES] ========== FUNCTION ENTRY ==========");
    } catch (initErr) {
        return "Error at init: " + initErr;
    }

    try {
        MDUX_debugLog("[RESET-STROKES] Checking for active document...");
        var doc = null;
        try {
            doc = app.activeDocument;
        } catch (docErr) {
            MDUX_debugLog("[RESET-STROKES] Error getting activeDocument: " + docErr);
            return "Error: Cannot access active document";
        }

        if (!doc) {
            MDUX_debugLog("[RESET-STROKES] No document open");
            return "No document open";
        }
        MDUX_debugLog("[RESET-STROKES] Document: " + doc.name);

        MDUX_debugLog("[RESET-STROKES] Getting selection...");
        var sel = null;
        try {
            sel = doc.selection;
        } catch (selErr) {
            MDUX_debugLog("[RESET-STROKES] Error getting selection: " + selErr);
            return "Error: Cannot access selection";
        }

        if (!sel || sel.length === 0) {
            MDUX_debugLog("[RESET-STROKES] No selection");
            return "No selection";
        }
        MDUX_debugLog("[RESET-STROKES] Selection has " + sel.length + " items");

        // Layer to style mappings - ductwork lines
        var LAYER_STYLE_MAP = {
            "Green Ductwork": "Green Ductwork",
            "Light Green Ductwork": "Light Green Ductwork",
            "Blue Ductwork": "Blue Ductwork",
            "Orange Ductwork": "Orange Ductwork",
            "Light Orange Ductwork": "Light Orange Ductwork",
            "Thermostat Lines": "Thermostat Lines"
        };

        var count = 0;
        var errors = [];
        var styleCache = {};

        // Get graphic style by name
        function getStyle(styleName) {
            if (styleCache[styleName] !== undefined) return styleCache[styleName];
            try {
                MDUX_debugLog("[RESET-STROKES] Looking for style: " + styleName);
                var styleCount = doc.graphicStyles.length;
                MDUX_debugLog("[RESET-STROKES] Document has " + styleCount + " graphic styles");
                for (var i = 0; i < styleCount; i++) {
                    try {
                        var gs = doc.graphicStyles[i];
                        if (gs && gs.name === styleName) {
                            styleCache[styleName] = gs;
                            MDUX_debugLog("[RESET-STROKES] Found style: " + styleName);
                            return gs;
                        }
                    } catch (styleItemErr) {
                        MDUX_debugLog("[RESET-STROKES] Error accessing style " + i + ": " + styleItemErr);
                    }
                }
            } catch (e) {
                MDUX_debugLog("[RESET-STROKES] Error in getStyle: " + e);
            }
            MDUX_debugLog("[RESET-STROKES] Style NOT found: " + styleName);
            styleCache[styleName] = null;
            return null;
        }

        function applyStyleToItem(item) {
            try {
                var layerName = null;
                try {
                    layerName = item.layer ? item.layer.name : null;
                } catch (layerErr) {
                    MDUX_debugLog("[RESET-STROKES] Error getting layer: " + layerErr);
                    return;
                }

                MDUX_debugLog("[RESET-STROKES] Item type=" + item.typename + ", layer=" + layerName);
                if (!layerName) return;

                var styleName = LAYER_STYLE_MAP[layerName];
                if (!styleName) {
                    MDUX_debugLog("[RESET-STROKES] No style mapping for layer: " + layerName);
                    return;
                }

                var style = getStyle(styleName);
                if (!style) {
                    MDUX_debugLog("[RESET-STROKES] Style not available: " + styleName);
                    return;
                }

                MDUX_debugLog("[RESET-STROKES] Applying style " + styleName + " to item...");
                try {
                    style.applyTo(item);
                    MDUX_debugLog("[RESET-STROKES] Applied style successfully");
                    count++;
                } catch (applyErr) {
                    MDUX_debugLog("[RESET-STROKES] Error applying style: " + applyErr);
                    errors.push("Apply error: " + applyErr);
                }
            } catch (e) {
                MDUX_debugLog("[RESET-STROKES] Error in applyStyleToItem: " + e);
                errors.push("Item error: " + e);
            }
        }

        function walkItems(item) {
            if (!item) return;
            try {
                var typeName = item.typename;
                MDUX_debugLog("[RESET-STROKES] Walking item type: " + typeName);

                if (typeName === "GroupItem") {
                    try {
                        var itemCount = item.pageItems.length;
                        for (var i = 0; i < itemCount; i++) {
                            walkItems(item.pageItems[i]);
                        }
                    } catch (groupErr) {
                        MDUX_debugLog("[RESET-STROKES] Error walking group: " + groupErr);
                    }
                } else if (typeName === "CompoundPathItem") {
                    applyStyleToItem(item);
                } else if (typeName === "PathItem") {
                    try {
                        if (!item.guides && !item.clipping) {
                            applyStyleToItem(item);
                        }
                    } catch (pathErr) {
                        MDUX_debugLog("[RESET-STROKES] Error checking path properties: " + pathErr);
                    }
                }
            } catch (walkErr) {
                MDUX_debugLog("[RESET-STROKES] Error in walkItems: " + walkErr);
            }
        }

        MDUX_debugLog("[RESET-STROKES] Starting to walk " + sel.length + " selected items...");
        for (var i = 0; i < sel.length; i++) {
            try {
                walkItems(sel[i]);
            } catch (loopErr) {
                MDUX_debugLog("[RESET-STROKES] Error processing item " + i + ": " + loopErr);
            }
        }

        MDUX_debugLog("[RESET-STROKES] ========== COMPLETE ==========");
        MDUX_debugLog("[RESET-STROKES] Applied styles to " + count + " items, " + errors.length + " errors");

        if (errors.length > 0) {
            return "Applied to " + count + " item(s), " + errors.length + " error(s)";
        }
        return "Applied styles to " + count + " item(s)";
    } catch (e) {
        try {
            MDUX_debugLog("[RESET-STROKES] FATAL ERROR: " + e + " (line: " + e.line + ")");
        } catch (logErr) {}
        return "Error: " + (e.message || e.toString());
    }
}

// --- RESET DUCTWORK PARTS SCALE: Reset ductwork parts to original scale while keeping rotation ---
function MDUX_resetDuctworkPartsScale() {
    // Outer safety wrapper
    try {
        if (typeof MDUX_debugLog !== "function") {
            return "Error: MDUX_debugLog not defined";
        }
        MDUX_debugLog("[RESET-PARTS-SCALE] ========== FUNCTION ENTRY ==========");
    } catch (initErr) {
        return "Error at init: " + initErr;
    }

    try {
        MDUX_debugLog("[RESET-PARTS-SCALE] Checking for active document...");
        var doc = null;
        try {
            doc = app.activeDocument;
        } catch (docErr) {
            MDUX_debugLog("[RESET-PARTS-SCALE] Error getting activeDocument: " + docErr);
            return "Error: Cannot access active document";
        }

        if (!doc) {
            MDUX_debugLog("[RESET-PARTS-SCALE] No document open");
            return "No document open";
        }
        MDUX_debugLog("[RESET-PARTS-SCALE] Document: " + doc.name);

        MDUX_debugLog("[RESET-PARTS-SCALE] Getting selection...");
        var sel = null;
        try {
            sel = doc.selection;
        } catch (selErr) {
            MDUX_debugLog("[RESET-PARTS-SCALE] Error getting selection: " + selErr);
            return "Error: Cannot access selection";
        }

        if (!sel || sel.length === 0) {
            MDUX_debugLog("[RESET-PARTS-SCALE] No selection");
            return "No selection";
        }
        MDUX_debugLog("[RESET-PARTS-SCALE] Selection has " + sel.length + " items");

        var count = 0;
        var skipped = 0;
        var errors = [];

        function resetPartScale(item) {
            try {
                MDUX_debugLog("[RESET-PARTS-SCALE] Processing " + item.typename);

                // Get current scale from metadata
                var currentScaleTag = null;
                try {
                    currentScaleTag = MDUX_getTag(item, "MDUX_CurrentScale");
                } catch (tagErr) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Error getting tag: " + tagErr);
                    skipped++;
                    return;
                }
                MDUX_debugLog("[RESET-PARTS-SCALE] MDUX_CurrentScale tag = " + currentScaleTag);

                if (currentScaleTag === null) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] No scale metadata, skipping");
                    skipped++;
                    return;
                }

                var currentScale = parseFloat(currentScaleTag);
                if (isNaN(currentScale) || currentScale <= 0) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Invalid scale value: " + currentScaleTag);
                    skipped++;
                    return;
                }

                // If already at 100%, skip
                if (Math.abs(currentScale - 100) < 1) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Already at 100%, skipping");
                    return;
                }

                // Calculate the scale factor needed to return to 100%
                // If current is 50%, we need to scale by 200% (100/50 * 100)
                var resetScalePercent = (100 / currentScale) * 100;
                MDUX_debugLog("[RESET-PARTS-SCALE] Current: " + currentScale + "%, reset factor: " + resetScalePercent + "%");

                // Apply uniform scale to reset to original size
                try {
                    var matrix = app.getScaleMatrix(resetScalePercent, resetScalePercent);
                    // Use simpler transform call that works with more item types
                    item.transform(matrix);
                    MDUX_debugLog("[RESET-PARTS-SCALE] Transform applied successfully");
                } catch (transformErr) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Transform error: " + transformErr);
                    errors.push("Transform error: " + transformErr);
                    skipped++;
                    return;
                }

                // Update metadata
                try {
                    MDUX_setTag(item, "MDUX_CurrentScale", "100");
                    MDUX_debugLog("[RESET-PARTS-SCALE] Set MDUX_CurrentScale to 100");
                } catch (setTagErr) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Error setting tag: " + setTagErr);
                }

                count++;
            } catch (e) {
                MDUX_debugLog("[RESET-PARTS-SCALE] Error in resetPartScale: " + e);
                errors.push("Reset error: " + e);
                skipped++;
            }
        }

        function walkItems(item) {
            if (!item) return;
            try {
                var typeName = item.typename;
                MDUX_debugLog("[RESET-PARTS-SCALE] Walking item type: " + typeName);

                // Check if item has MDUX_CurrentScale metadata
                var hasScaleTag = false;
                try {
                    hasScaleTag = MDUX_getTag(item, "MDUX_CurrentScale") !== null;
                } catch (tagCheckErr) {
                    MDUX_debugLog("[RESET-PARTS-SCALE] Error checking tag: " + tagCheckErr);
                }

                // Handle PlacedItem (linked/embedded images or symbols)
                if (typeName === "PlacedItem" || typeName === "SymbolItem") {
                    if (hasScaleTag) {
                        resetPartScale(item);
                    } else {
                        MDUX_debugLog("[RESET-PARTS-SCALE] " + typeName + " has no scale metadata");
                        skipped++;
                    }
                } else if (typeName === "GroupItem") {
                    // Groups with MDUX metadata are ductwork parts (Units, Registers, etc.)
                    if (hasScaleTag) {
                        resetPartScale(item);
                    } else {
                        // Walk into group to find nested items
                        try {
                            var itemCount = item.pageItems.length;
                            for (var i = 0; i < itemCount; i++) {
                                walkItems(item.pageItems[i]);
                            }
                        } catch (groupErr) {
                            MDUX_debugLog("[RESET-PARTS-SCALE] Error walking group: " + groupErr);
                        }
                    }
                } else if (typeName === "CompoundPathItem" || typeName === "PathItem") {
                    // Check if this item has scale metadata stored
                    if (hasScaleTag) {
                        resetPartScale(item);
                    }
                }
            } catch (walkErr) {
                MDUX_debugLog("[RESET-PARTS-SCALE] Error in walkItems: " + walkErr);
            }
        }

        MDUX_debugLog("[RESET-PARTS-SCALE] Starting to walk " + sel.length + " selected items...");
        for (var i = 0; i < sel.length; i++) {
            try {
                walkItems(sel[i]);
            } catch (loopErr) {
                MDUX_debugLog("[RESET-PARTS-SCALE] Error processing item " + i + ": " + loopErr);
            }
        }

        MDUX_debugLog("[RESET-PARTS-SCALE] ========== COMPLETE ==========");
        MDUX_debugLog("[RESET-PARTS-SCALE] Reset " + count + " parts, skipped " + skipped + ", errors " + errors.length);

        if (count === 0 && skipped > 0) {
            return "No parts with scale metadata found (checked " + skipped + " items)";
        }
        return "Reset scale on " + count + " part(s) to 100%" + (skipped > 0 ? ", skipped " + skipped : "");
    } catch (e) {
        try {
            MDUX_debugLog("[RESET-PARTS-SCALE] FATAL ERROR: " + e + " (line: " + e.line + ")");
        } catch (logErr) {}
        return "Error: " + e.message;
    }
}
