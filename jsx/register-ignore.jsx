if (typeof MDUX === "undefined") {
    var MDUX = {};
}

MDUX.previewLayerName = "__MDUX_RegisterPreview";
MDUX.ignoreLayerNames = ["Ignore", "Ignored", "ignore", "ignored"];
MDUX.ignoreModeActive = false;
MDUX.selectionListener = null;
MDUX.currentPreviewInfo = null;
MDUX.previewStrokeColor = (function () {
    var c = new RGBColor();
    c.red = 255;
    c.green = 102;
    c.blue = 0;
    return c;
})();
MDUX.previewFillColor = (function () {
    var c = new RGBColor();
    c.red = 255;
    c.green = 180;
    c.blue = 0;
    return c;
})();

function MDUX_log(msg) {
    try {
        $.writeln("[Magic-Ductwork] " + msg);
    } catch (e) {}
}

function MDUX_isIgnoreModeActive() {
    return !!MDUX.ignoreModeActive;
}

function MDUX_toggleIgnoreMode() {
    return MDUX_setIgnoreMode(!MDUX.ignoreModeActive);
}

function MDUX_setIgnoreMode(state) {
    state = !!state;
    if (state === MDUX.ignoreModeActive) {
        return state ? "true" : "false";
    }
    if (state) {
        MDUX_activateIgnoreMode();
    } else {
        MDUX_deactivateIgnoreMode();
    }
    MDUX.ignoreModeActive = state;
    return state ? "true" : "false";
}

function MDUX_activateIgnoreMode() {
    if (app.documents.length === 0) {
        throw new Error("No open document.");
    }
    MDUX_teardownListener();
    MDUX.selectionListener = function () {
        try {
            MDUX_refreshPreview();
        } catch (e) {
            MDUX_log("Selection listener error: " + e);
        }
    };
    app.addEventListener("afterSelectionChanged", MDUX.selectionListener);
    MDUX_refreshPreview();
}

function MDUX_deactivateIgnoreMode() {
    MDUX_teardownListener();
    if (app.documents.length > 0) {
        try {
            MDUX_clearPreview(app.activeDocument);
        } catch (e) {}
    }
    MDUX.currentPreviewInfo = null;
}

function MDUX_teardownListener() {
    if (MDUX.selectionListener) {
        try {
            app.removeEventListener("afterSelectionChanged", MDUX.selectionListener);
        } catch (e) {}
        MDUX.selectionListener = null;
    }
}

function MDUX_refreshPreview() {
    if (app.documents.length === 0) {
        MDUX_clearPreview(null);
        MDUX.currentPreviewInfo = null;
        return;
    }
    var doc = app.activeDocument;
    MDUX_clearPreview(doc);
    var sel = null;
    try {
        sel = doc.selection;
    } catch (eSel) {
        sel = null;
    }
    if (!sel || sel.length === 0) {
        MDUX.currentPreviewInfo = null;
        return;
    }
    var pathItem = null;
    for (var i = 0; i < sel.length; i++) {
        if (sel[i] instanceof PathItem) {
            if (!sel[i].closed && sel[i].pathPoints.length >= 2) {
                pathItem = sel[i];
                break;
            }
        }
    }
    if (!pathItem) {
        MDUX.currentPreviewInfo = null;
        return;
    }
    MDUX.currentPreviewInfo = MDUX_buildPreview(doc, pathItem);
}

function MDUX_buildPreview(doc, pathItem) {
    var previewLayer = MDUX_getPreviewLayer(doc, true);
    if (!previewLayer) {
        return null;
    }
    var clone = pathItem.duplicate(previewLayer, ElementPlacement.PLACEATEND);
    clone.name = "__MDUX_preview_path";
    clone.filled = false;
    clone.stroked = true;
    clone.strokeColor = MDUX.previewStrokeColor;
    try { clone.strokeWidth = Math.max(1.5, pathItem.strokeWidth || 1); } catch (eSW) { clone.strokeWidth = 1.5; }
    clone.opacity = 50;

    var info = MDUX_calculateEndpointInfo(pathItem);
    if (!info) {
        return null;
    }

    var dotRadius = info.length / 6;
    if (dotRadius < 3) dotRadius = 3;
    if (dotRadius > 15) dotRadius = 15;

    var ellipseTop = info.endpoint[1] + dotRadius;
    var ellipseLeft = info.endpoint[0] - dotRadius;
    var dot = previewLayer.pathItems.ellipse(ellipseTop, ellipseLeft, dotRadius * 2, dotRadius * 2);
    dot.stroked = false;
    dot.filled = true;
    dot.fillColor = MDUX.previewFillColor;
    dot.opacity = 80;
    dot.name = "__MDUX_preview_dot";

    var arrow = previewLayer.pathItems.add();
    arrow.stroked = true;
    arrow.strokeColor = MDUX.previewStrokeColor;
    arrow.strokeWidth = 1.5;
    arrow.filled = false;
    arrow.closed = false;
    var tail = [info.endpoint[0] - info.direction[0] * dotRadius * 2,
                info.endpoint[1] - info.direction[1] * dotRadius * 2];
    arrow.setEntirePath([info.endpoint, tail]);
    arrow.opacity = 80;
    arrow.name = "__MDUX_preview_arrow";

    return {
        document: doc,
        path: pathItem,
        endpoint: info.endpoint,
        direction: info.direction,
        length: info.length
    };
}

function MDUX_calculateEndpointInfo(pathItem) {
    if (!pathItem || pathItem.closed || pathItem.pathPoints.length < 2) {
        return null;
    }
    var pts = pathItem.pathPoints;
    var last = pts[pts.length - 1];
    var prev = pts[pts.length - 2];
    var endpoint = [last.anchor[0], last.anchor[1]];
    var dx = endpoint[0] - prev.anchor[0];
    var dy = endpoint[1] - prev.anchor[1];
    var length = Math.sqrt(dx * dx + dy * dy);
    if (length === 0) {
        dx = 1;
        dy = 0;
        length = 1;
    }
    var dir = [dx / length, dy / length];
    return {
        endpoint: endpoint,
        direction: dir,
        length: length
    };
}

function MDUX_clearPreview(doc) {
    if (!doc) {
        return;
    }
    var layer = MDUX_getPreviewLayer(doc, false);
    if (!layer) return;
    try {
        while (layer.pageItems.length) {
            layer.pageItems[0].remove();
        }
    } catch (e) {}
}

function MDUX_getPreviewLayer(doc, createIfMissing) {
    if (!doc) return null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === MDUX.previewLayerName) {
            doc.layers[i].locked = false;
            doc.layers[i].visible = true;
            return doc.layers[i];
        }
    }
    if (!createIfMissing) return null;
    var layer = doc.layers.add();
    layer.name = MDUX.previewLayerName;
    layer.zOrder(ZOrderMethod.SENDTOFRONT);
    layer.locked = false;
    layer.visible = true;
    return layer;
}

function MDUX_getIgnoreLayer(doc, createIfMissing) {
    if (!doc) return null;
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        var candidate = doc.layers[i];
        for (var j = 0; j < MDUX.ignoreLayerNames.length; j++) {
            if (candidate.name === MDUX.ignoreLayerNames[j]) {
                layer = candidate;
                break;
            }
        }
        if (layer) break;
    }
    if (!layer && createIfMissing) {
        layer = doc.layers.add();
        layer.name = "Ignore";
        layer.zOrder(ZOrderMethod.SENDTOFRONT);
    }
    if (layer) {
        try { layer.locked = false; } catch (e) {}
        try { layer.visible = true; } catch (eVis) {}
    }
    return layer;
}

function MDUX_applyIgnoreToCurrent() {
    if (!MDUX.currentPreviewInfo || !MDUX.currentPreviewInfo.path) {
        return "NO_SELECTION";
    }
    var doc = null;
    try {
        doc = MDUX.currentPreviewInfo.document || app.activeDocument;
    } catch (eDoc) {
        doc = app.activeDocument;
    }
    if (!doc) {
        return "NO_DOCUMENT";
    }
    if (!MDUX.currentPreviewInfo.path) {
        MDUX_refreshPreview();
        return "INVALID_SELECTION";
    }
    var ignoreLayer = MDUX_getIgnoreLayer(doc, true);
    if (!ignoreLayer) {
        return "NO_IGNORE_LAYER";
    }
    var pt = MDUX.currentPreviewInfo.endpoint;
    var ignorePath = ignoreLayer.pathItems.add();
    ignorePath.stroked = false;
    ignorePath.filled = false;
    ignorePath.closed = false;
    ignorePath.name = "__MDUX_ignore_point";
    ignorePath.setEntirePath([pt, pt]);
    ignorePath.opacity = 0;

    MDUX_refreshPreview();
    return "IGNORE_ADDED";
}

function MDUX_getIgnoreModeStatus() {
    if (!MDUX.ignoreModeActive) {
        return "inactive";
    }
    if (!MDUX.currentPreviewInfo) {
        return "active:no-selection";
    }
    try {
        var name = MDUX.currentPreviewInfo.path.name || "Path";
        return "active:" + name;
    } catch (e) {
        return "active";
    }
}

function MDUX_cleanup() {
    MDUX_deactivateIgnoreMode();
}
