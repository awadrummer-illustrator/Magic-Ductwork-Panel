// Gap tools for manual heal/recreate workflows

function MDUX_gapTools_isExcludedLayerName(name) {
    if (!name) return false;
    return name === "Gap Definitions" ||
        name === "Deleted Segments" ||
        name === "Ignore" ||
        name === "Ignored" ||
        name === "ignore" ||
        name === "ignored";
}

function MDUX_gapTools_getParentLayer(item) {
    var cur = item;
    while (cur && cur.typename !== "Layer") {
        try {
            cur = cur.parent;
        } catch (e) {
            cur = null;
        }
    }
    return cur || null;
}

function MDUX_gapTools_collectSelectedPaths(doc) {
    var results = [];

    function addPath(path) {
        if (!path) return;
        try { if (path.isValid === false) return; } catch (e) { }
        try { if (path.guides) return; } catch (e) { }
        try { if (path.clipping) return; } catch (e) { }
        try {
            if (path.layer && MDUX_gapTools_isExcludedLayerName(path.layer.name)) return;
        } catch (e) { }
        try { if (path.locked || path.hidden) return; } catch (e) { }
        try { if (!path.pathPoints || path.pathPoints.length < 2) return; } catch (e) { return; }

        for (var i = 0; i < results.length; i++) {
            if (results[i] === path) return;
        }
        results.push(path);
    }

    function collect(item) {
        if (!item) return;
        var typeName = "";
        try { typeName = item.typename; } catch (e) { typeName = ""; }

        if (typeName === "PathItem") {
            addPath(item);
            return;
        }
        if (typeName === "CompoundPathItem") {
            try {
                for (var i = 0; i < item.pathItems.length; i++) {
                    addPath(item.pathItems[i]);
                }
            } catch (e) { }
            return;
        }

        try {
            if (item.pageItems && item.pageItems.length !== undefined) {
                for (var p = 0; p < item.pageItems.length; p++) {
                    collect(item.pageItems[p]);
                }
            }
        } catch (e) { }

        try {
            if (item.groupItems && item.groupItems.length !== undefined) {
                for (var g = 0; g < item.groupItems.length; g++) {
                    collect(item.groupItems[g]);
                }
            }
        } catch (e) { }

        try {
            if (item.pathItems && !item.pageItems && typeName !== "CompoundPathItem") {
                for (var pi = 0; pi < item.pathItems.length; pi++) {
                    addPath(item.pathItems[pi]);
                }
            }
        } catch (e) { }
    }

    var sel = null;
    try { sel = doc.selection; } catch (e) { sel = null; }
    if (!sel) return results;

    if (sel.length === undefined && sel.typename) {
        collect(sel);
        return results;
    }

    for (var s = 0; s < sel.length; s++) {
        collect(sel[s]);
    }

    return results;
}

function MDUX_gapTools_getSelectionBounds(paths) {
    if (!paths || paths.length < 1) return null;
    var minX = 1e9, minY = 1e9, maxX = -1e9, maxY = -1e9;
    var valid = false;

    for (var i = 0; i < paths.length; i++) {
        try {
            var b = paths[i].geometricBounds;
            if (!b || b.length < 4) continue;
            var left = b[0], top = b[1], right = b[2], bottom = b[3];
            minX = Math.min(minX, left, right);
            maxX = Math.max(maxX, left, right);
            minY = Math.min(minY, top, bottom);
            maxY = Math.max(maxY, top, bottom);
            valid = true;
        } catch (e) { }
    }

    if (!valid) return null;
    return { minX: minX, minY: minY, maxX: maxX, maxY: maxY };
}

function MDUX_gapTools_pointDistance(a, b) {
    var dx = a[0] - b[0];
    var dy = a[1] - b[1];
    return Math.sqrt(dx * dx + dy * dy);
}

function MDUX_gapTools_pointInBounds(pt, bounds, margin) {
    if (!bounds) return true;
    var m = (typeof margin === "number") ? margin : 0;
    return pt[0] >= bounds.minX - m && pt[0] <= bounds.maxX + m &&
        pt[1] >= bounds.minY - m && pt[1] <= bounds.maxY + m;
}

function MDUX_gapTools_closestPointOnSegment(pt, a, b) {
    var ax = a[0], ay = a[1];
    var bx = b[0], by = b[1];
    var dx = bx - ax;
    var dy = by - ay;
    var len2 = dx * dx + dy * dy;
    if (len2 === 0) {
        return { pt: [ax, ay], t: 0, dist: MDUX_gapTools_pointDistance(pt, a), segLen: 0 };
    }
    var t = ((pt[0] - ax) * dx + (pt[1] - ay) * dy) / len2;
    if (t < 0) t = 0;
    if (t > 1) t = 1;
    var proj = [ax + t * dx, ay + t * dy];
    var dist = MDUX_gapTools_pointDistance(pt, proj);
    var segLen = Math.sqrt(len2);
    return { pt: proj, t: t, dist: dist, segLen: segLen };
}

function MDUX_gapTools_findNearestSegment(paths, point, layerName, maxDist) {
    if (!paths || paths.length < 1) return null;
    var best = null;
    var bestDist = (typeof maxDist === "number") ? maxDist : 1e9;

    for (var i = 0; i < paths.length; i++) {
        var path = paths[i];
        try { if (path.isValid === false) continue; } catch (e) { }
        try {
            if (layerName && path.layer && path.layer.name !== layerName) continue;
        } catch (e) { }
        var pts = null;
        try { pts = path.pathPoints; } catch (e) { pts = null; }
        if (!pts || pts.length < 2) continue;

        for (var s = 0; s < pts.length - 1; s++) {
            var a = pts[s].anchor;
            var b = pts[s + 1].anchor;
            var res = MDUX_gapTools_closestPointOnSegment(point, a, b);
            if (res.segLen < 0.01) continue;
            if (res.dist <= bestDist) {
                bestDist = res.dist;
                best = {
                    path: path,
                    segIdx: s,
                    proj: res.pt,
                    dist: res.dist,
                    segStart: [a[0], a[1]],
                    segEnd: [b[0], b[1]],
                    segLen: res.segLen,
                    t: res.t
                };
            }
        }
    }
    return best;
}

function MDUX_gapTools_parseGapMetadata(note) {
    if (!note || note.indexOf("MDUX_GAP:") !== 0) return null;
    try { return JSON.parse(note.substring(9)); } catch (e) { return null; }
}

function MDUX_gapTools_getMarkerInfo(marker) {
    var meta = null;
    try { meta = MDUX_gapTools_parseGapMetadata(marker.note); } catch (e) { meta = null; }
    var center = null;
    var gapSize = null;
    var dir = null;
    var sourceLayer = null;
    var isAutoSized = true;

    if (meta) {
        if (typeof meta.x === "number" && typeof meta.y === "number") center = [meta.x, meta.y];
        if (typeof meta.gapSize === "number") gapSize = meta.gapSize;
        if (typeof meta.dirX === "number" && typeof meta.dirY === "number") dir = [meta.dirX, meta.dirY];
        if (typeof meta.sourceLayer === "string") sourceLayer = meta.sourceLayer;
        if (meta.isAutoSized === false) isAutoSized = false;
    }

    try {
        if (!center && marker.pathPoints && marker.pathPoints.length >= 2) {
            var mA = marker.pathPoints[0].anchor;
            var mB = marker.pathPoints[marker.pathPoints.length - 1].anchor;
            center = [(mA[0] + mB[0]) / 2, (mA[1] + mB[1]) / 2];
        }
    } catch (e) { }

    try {
        if (!gapSize && marker.pathPoints && marker.pathPoints.length >= 2) {
            var mA2 = marker.pathPoints[0].anchor;
            var mB2 = marker.pathPoints[marker.pathPoints.length - 1].anchor;
            var len = MDUX_gapTools_pointDistance(mA2, mB2);
            if (len > 0) gapSize = len / 2;
        }
    } catch (e) { }

    try {
        if (!dir && marker.pathPoints && marker.pathPoints.length >= 2) {
            var mA3 = marker.pathPoints[0].anchor;
            var mB3 = marker.pathPoints[marker.pathPoints.length - 1].anchor;
            var dx = mB3[0] - mA3[0];
            var dy = mB3[1] - mA3[1];
            var lenDir = Math.sqrt(dx * dx + dy * dy);
            if (lenDir > 0.001) {
                dir = [dx / lenDir, dy / lenDir];
            }
        }
    } catch (e) { }

    if (!dir) dir = [1, 0];

    return {
        center: center,
        gapSize: gapSize,
        dir: dir,
        sourceLayer: sourceLayer,
        isAutoSized: isAutoSized,
        meta: meta
    };
}

function MDUX_gapTools_getGapLayer(doc, createIfMissing) {
    var layer = null;
    try { layer = doc.layers.getByName("Gap Definitions"); } catch (e) { layer = null; }
    if (!layer && createIfMissing) {
        try {
            layer = doc.layers.add();
            layer.name = "Gap Definitions";
        } catch (e) { layer = null; }
    }
    return layer;
}

function MDUX_gapTools_findDeletedSegmentNear(layer, center, tol) {
    if (!layer || !center) return null;
    var best = null;
    var bestDist = (typeof tol === "number") ? tol : 9999;
    try {
        for (var i = 0; i < layer.pathItems.length; i++) {
            var seg = layer.pathItems[i];
            if (!seg || !seg.pathPoints || seg.pathPoints.length < 2) continue;
            var a = seg.pathPoints[0].anchor;
            var b = seg.pathPoints[seg.pathPoints.length - 1].anchor;
            var mid = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
            var dist = MDUX_gapTools_pointDistance(center, mid);
            if (dist <= bestDist) {
                bestDist = dist;
                best = seg;
            }
        }
    } catch (e) { }
    return best;
}

function MDUX_gapTools_restoreDeletedSegment(seg, targetLayer) {
    if (!seg || !targetLayer) return null;
    var newPath = null;
    try {
        var pts = seg.pathPoints;
        if (!pts || pts.length < 2) return null;
        var start = [pts[0].anchor[0], pts[0].anchor[1]];
        var end = [pts[pts.length - 1].anchor[0], pts[pts.length - 1].anchor[1]];
        newPath = targetLayer.pathItems.add();
        newPath.setEntirePath([start, end]);
        newPath.filled = false;
        newPath.stroked = true;
        try { newPath.strokeWidth = seg.strokeWidth; } catch (e) { }
        try { newPath.strokeColor = seg.strokeColor; } catch (e) { }
        try { seg.remove(); } catch (e) { }
    } catch (e) {
        newPath = null;
    }
    return newPath;
}

function MDUX_gapTools_findPatchMarkerNear(gapLayer, center, tol) {
    if (!gapLayer || !center) return null;
    var best = null;
    var bestDist = (typeof tol === "number") ? tol : 9999;
    for (var i = 0; i < gapLayer.pathItems.length; i++) {
        try {
            var item = gapLayer.pathItems[i];
            if (!item.note || item.note.indexOf("MDUX_PATCH") !== 0) continue;
            if (!item.pathPoints || item.pathPoints.length < 1) continue;
            var pt = item.pathPoints[0].anchor;
            var dist = MDUX_gapTools_pointDistance(center, pt);
            if (dist <= bestDist) {
                bestDist = dist;
                best = item;
            }
        } catch (e) { }
    }
    return best;
}

function MDUX_gapTools_createPatchMarker(gapLayer, center) {
    if (!gapLayer || !center) return null;
    var existing = MDUX_gapTools_findPatchMarkerNear(gapLayer, center, 6);
    if (existing) return existing;
    try {
        var marker = gapLayer.pathItems.add();
        marker.setEntirePath([center, center]);
        marker.stroked = false;
        marker.filled = false;
        marker.note = "MDUX_PATCH";
        marker.locked = true;
        return marker;
    } catch (e) { return null; }
}

function MDUX_gapTools_removeGapDefinitionAnchorsNear(gapLayer, center, tol) {
    if (!gapLayer || !center) return 0;
    var removed = 0;
    var t = (typeof tol === "number") ? tol : 8;
    for (var i = gapLayer.pathItems.length - 1; i >= 0; i--) {
        try {
            var item = gapLayer.pathItems[i];
            if (!item) continue;
            if (item.note && (item.note.indexOf("MDUX_GAP:") === 0 || item.note.indexOf("MDUX_PATCH") === 0)) {
                continue;
            }
            if (!item.pathPoints || item.pathPoints.length < 1) continue;
            var pt = item.pathPoints[0].anchor;
            if (MDUX_gapTools_pointDistance(center, pt) <= t) {
                item.remove();
                removed++;
            }
        } catch (e) { }
    }
    return removed;
}

function MDUX_gapTools_getIgnoreLayer(doc, createIfMissing) {
    var names = ["Ignored", "Ignore", "ignored", "ignore"];
    var layer = null;
    for (var i = 0; i < names.length; i++) {
        try {
            layer = doc.layers.getByName(names[i]);
            if (layer) break;
        } catch (e) { }
    }
    if (!layer && createIfMissing) {
        try {
            layer = doc.layers.add();
            layer.name = "Ignored";
        } catch (e) { layer = null; }
    }
    return layer;
}

function MDUX_gapTools_isIgnoreAnchorNear(ignoreLayer, pt, tol) {
    if (!ignoreLayer || !pt) return false;
    var t = (typeof tol === "number") ? tol : 4;
    try {
        for (var i = 0; i < ignoreLayer.pathItems.length; i++) {
            var item = ignoreLayer.pathItems[i];
            if (!item || !item.pathPoints || item.pathPoints.length < 1) continue;
            var anchor = item.pathPoints[0].anchor;
            if (MDUX_gapTools_pointDistance(anchor, pt) <= t) return true;
        }
    } catch (e) { }
    return false;
}

function MDUX_gapTools_createIgnoreAnchor(ignoreLayer, pt) {
    if (!ignoreLayer || !pt) return false;
    if (MDUX_gapTools_isIgnoreAnchorNear(ignoreLayer, pt, 4)) return false;
    try {
        var p = ignoreLayer.pathItems.add();
        p.setEntirePath([pt, pt]);
        p.stroked = false;
        p.filled = false;
        p.name = "__MDUX_ignore_point";
        p.opacity = 0;
        return true;
    } catch (e) { }
    return false;
}

function MDUX_gapTools_removeIgnoreAnchorsNear(doc, points, tol) {
    if (!doc || !points || points.length < 1) return 0;
    var removed = 0;
    var layer = MDUX_gapTools_getIgnoreLayer(doc, false);
    if (!layer) return 0;

    var wasLocked = layer.locked;
    var wasVisible = layer.visible;
    try { if (wasLocked) layer.locked = false; } catch (e) { }
    try { if (!wasVisible) layer.visible = true; } catch (e) { }

    var t = (typeof tol === "number") ? tol : 10;
    for (var i = layer.pathItems.length - 1; i >= 0; i--) {
        try {
            var item = layer.pathItems[i];
            if (!item || !item.pathPoints || item.pathPoints.length < 1) continue;
            var anchor = item.pathPoints[0].anchor;
            for (var p = 0; p < points.length; p++) {
                if (MDUX_gapTools_pointDistance(anchor, points[p]) <= t) {
                    item.remove();
                    removed++;
                    break;
                }
            }
        } catch (e) { }
    }

    try { layer.locked = wasLocked; } catch (e) { }
    try { layer.visible = wasVisible; } catch (e) { }
    return removed;
}

function MDUX_gapTools_calculateAutoGapSize(strokeWidth) {
    var base = (strokeWidth && strokeWidth > 0) ? (strokeWidth * 0.6) : 4.25;
    if (base < 4) base = 4;
    return base;
}

function MDUX_gapTools_createGapMarker(doc, gapLayer, center, gapSize, sourceLayer, dir, isAutoSized, existingMarker) {
    if (!doc || !gapLayer || !center || !gapSize || !dir) return null;

    var marker = existingMarker || null;
    if (!marker) {
        for (var i = 0; i < gapLayer.pathItems.length; i++) {
            try {
                var chk = gapLayer.pathItems[i];
                if (!chk.note || chk.note.indexOf("MDUX_GAP:") !== 0) continue;
                var meta = MDUX_gapTools_parseGapMetadata(chk.note);
                if (!meta || typeof meta.x !== "number" || typeof meta.y !== "number") continue;
                var dist = MDUX_gapTools_pointDistance(center, [meta.x, meta.y]);
                if (dist < 10) { marker = chk; break; }
            } catch (e) { }
        }
    }

    var wasLocked = gapLayer.locked;
    var wasVisible = gapLayer.visible;
    try { if (wasLocked) gapLayer.locked = false; } catch (e) { }
    try { if (!wasVisible) gapLayer.visible = true; } catch (e) { }

    try {
        if (!marker) {
            marker = gapLayer.pathItems.add();
        }
        try { marker.locked = false; } catch (e) { }

        var startPt = [center[0] - gapSize * dir[0], center[1] - gapSize * dir[1]];
        var endPt = [center[0] + gapSize * dir[0], center[1] + gapSize * dir[1]];
        marker.setEntirePath([startPt, endPt]);
        marker.filled = false;
        marker.stroked = true;
        marker.strokeWidth = 3;
        try {
            var magenta = new RGBColor();
            magenta.red = 255;
            magenta.green = 0;
            magenta.blue = 255;
            marker.strokeColor = magenta;
        } catch (e) { }

        var metadata = {
            x: center[0],
            y: center[1],
            gapSize: gapSize,
            sourceLayer: sourceLayer || "",
            isAutoSized: isAutoSized !== false,
            dirX: dir[0],
            dirY: dir[1],
            createdAt: new Date().toString()
        };
        marker.note = "MDUX_GAP:" + JSON.stringify(metadata);
        marker.locked = true;
    } catch (e) {
        marker = null;
    }

    try { gapLayer.locked = wasLocked; } catch (e) { }
    try { gapLayer.visible = wasVisible; } catch (e) { }

    return marker;
}

function MDUX_gapTools_saveDeletedSegment(doc, startPt, endPt, sourceLayer) {
    if (!doc || !startPt || !endPt) return null;
    var deletedLayer = null;
    try { deletedLayer = doc.layers.getByName("Deleted Segments"); } catch (e) { deletedLayer = null; }
    if (!deletedLayer) {
        try {
            deletedLayer = doc.layers.add();
            deletedLayer.name = "Deleted Segments";
        } catch (e) { deletedLayer = null; }
    }
    if (!deletedLayer) return null;

    var wasLocked = deletedLayer.locked;
    var wasVisible = deletedLayer.visible;
    try { if (wasLocked) deletedLayer.locked = false; } catch (e) { }
    try { if (!wasVisible) deletedLayer.visible = true; } catch (e) { }

    try {
        var center = [(startPt[0] + endPt[0]) / 2, (startPt[1] + endPt[1]) / 2];
        for (var i = deletedLayer.pathItems.length - 1; i >= 0; i--) {
            var ex = deletedLayer.pathItems[i];
            if (!ex.pathPoints || ex.pathPoints.length < 2) continue;
            var exA = ex.pathPoints[0].anchor;
            var exB = ex.pathPoints[ex.pathPoints.length - 1].anchor;
            var exCenter = [(exA[0] + exB[0]) / 2, (exA[1] + exB[1]) / 2];
            if (MDUX_gapTools_pointDistance(center, exCenter) < 10) {
                try { ex.remove(); } catch (e) { }
            }
        }
    } catch (e) { }

    var segPath = null;
    try {
        segPath = deletedLayer.pathItems.add();
        segPath.setEntirePath([[startPt[0], startPt[1]], [endPt[0], endPt[1]]]);
        segPath.filled = false;
        segPath.stroked = true;
        segPath.strokeWidth = 1;
        try {
            if (sourceLayer) {
                var src = null;
                try { src = doc.layers.getByName(sourceLayer); } catch (e) { src = null; }
                if (src && src.pathItems.length > 0) {
                    segPath.strokeColor = src.pathItems[0].strokeColor;
                }
            }
        } catch (e) { }
    } catch (e) { segPath = null; }

    try { deletedLayer.locked = wasLocked; } catch (e) { }
    try { deletedLayer.visible = wasVisible; } catch (e) { }

    return segPath;
}

function MDUX_gapTools_splitPathAtGap(path, segIdx, cutBefore, cutAfter, parentLayer) {
    if (!path || !cutBefore || !cutAfter) return null;
    var first = null, second = null;
    var targetLayer = parentLayer || null;
    if (!targetLayer) {
        try { targetLayer = MDUX_gapTools_getParentLayer(path); } catch (e) { targetLayer = null; }
    }
    try { if (!targetLayer) targetLayer = path.layer; } catch (e) { }

    try { first = path.duplicate(targetLayer, ElementPlacement.PLACEATEND); } catch (e) { first = null; }
    try { second = path.duplicate(targetLayer, ElementPlacement.PLACEATEND); } catch (e) { second = null; }
    if (!first || !second) return null;

    try {
        var pts1 = first.pathPoints;
        for (var i = pts1.length - 1; i > segIdx; i--) {
            try { pts1[i].remove(); } catch (e) { }
        }
        var newEnd = first.pathPoints.add();
        newEnd.anchor = cutBefore;
        newEnd.leftDirection = cutBefore;
        newEnd.rightDirection = cutBefore;
    } catch (e) { }

    try {
        var pts2 = second.pathPoints;
        if (segIdx < pts2.length) {
            pts2[segIdx].anchor = cutAfter;
            pts2[segIdx].leftDirection = cutAfter;
            pts2[segIdx].rightDirection = cutAfter;
        }
        for (var j = segIdx - 1; j >= 0; j--) {
            try { pts2[j].remove(); } catch (e) { }
        }
    } catch (e) { }

    try { path.remove(); } catch (e) { }
    return { first: first, second: second };
}

function MDUX_healGapsInSelection() {
    try {
        if (app.documents.length === 0) return "No document open";
        var doc = app.activeDocument;
        var selectedPaths = MDUX_gapTools_collectSelectedPaths(doc);
        if (!selectedPaths || selectedPaths.length < 1) return "Select ductwork paths to heal gaps";

        var bounds = MDUX_gapTools_getSelectionBounds(selectedPaths);
        var gapLayer = MDUX_gapTools_getGapLayer(doc, false);
        if (!gapLayer) return "No Gap Definitions layer";

        var deletedLayer = null;
        try { deletedLayer = doc.layers.getByName("Deleted Segments"); } catch (e) { deletedLayer = null; }

        var gapLayerWasLocked = gapLayer.locked;
        var gapLayerWasVisible = gapLayer.visible;
        try { if (gapLayerWasLocked) gapLayer.locked = false; } catch (e) { }
        try { if (!gapLayerWasVisible) gapLayer.visible = true; } catch (e) { }

        var restoredSegments = [];
        var healedCount = 0;
        var restoredCount = 0;
        var markerRemoved = 0;
        var patchedCount = 0;
        var ignoreRemoved = 0;
        var skippedMissing = 0;

        var markers = [];
        for (var i = 0; i < gapLayer.pathItems.length; i++) {
            var item = gapLayer.pathItems[i];
            if (!item || !item.note || item.note.indexOf("MDUX_GAP:") !== 0) continue;
            var info = MDUX_gapTools_getMarkerInfo(item);
            if (!info.center) continue;
            if (!MDUX_gapTools_pointInBounds(info.center, bounds, 25)) continue;
            var near = MDUX_gapTools_findNearestSegment(selectedPaths, info.center, info.sourceLayer, 25);
            if (!near) continue;
            markers.push({ marker: item, info: info });
        }

        for (var m = 0; m < markers.length; m++) {
            var entry = markers[m];
            var marker = entry.marker;
            var info = entry.info;
            var center = info.center;
            var gapSize = info.gapSize;
            var dir = info.dir;
            var sourceLayer = info.sourceLayer;

            var restored = null;
            if (deletedLayer) {
                try {
                    if (deletedLayer.locked) deletedLayer.locked = false;
                    if (!deletedLayer.visible) deletedLayer.visible = true;
                } catch (e) { }
                var seg = MDUX_gapTools_findDeletedSegmentNear(deletedLayer, center, 20);
                if (seg) {
                    var targetLayer = null;
                    try {
                        if (sourceLayer) targetLayer = doc.layers.getByName(sourceLayer);
                    } catch (e) { targetLayer = null; }
                    if (!targetLayer) {
                        var nearPath = MDUX_gapTools_findNearestSegment(selectedPaths, center, null, 25);
                        if (nearPath && nearPath.path && nearPath.path.layer) targetLayer = nearPath.path.layer;
                    }
                    if (targetLayer) {
                        restored = MDUX_gapTools_restoreDeletedSegment(seg, targetLayer);
                        if (restored) {
                            restoredCount++;
                            restoredSegments.push(restored);
                        }
                    }
                }
            }

            var filled = false;
            var nearFilled = MDUX_gapTools_findNearestSegment(selectedPaths, center, sourceLayer, 4);
            if (nearFilled && nearFilled.dist <= 4) filled = true;

            if (!restored && !filled) {
                skippedMissing++;
                continue;
            }

            var cutBefore = null;
            var cutAfter = null;
            if (gapSize && dir) {
                cutBefore = [center[0] - gapSize * dir[0], center[1] - gapSize * dir[1]];
                cutAfter = [center[0] + gapSize * dir[0], center[1] + gapSize * dir[1]];
            }

            if (cutBefore && cutAfter) {
                ignoreRemoved += MDUX_gapTools_removeIgnoreAnchorsNear(doc, [cutBefore, cutAfter], 12);
            }

            try {
                if (marker) {
                    try { marker.locked = false; } catch (e) { }
                    marker.remove();
                    markerRemoved++;
                }
            } catch (e) { }

            MDUX_gapTools_removeGapDefinitionAnchorsNear(gapLayer, center, 12);
            if (MDUX_gapTools_createPatchMarker(gapLayer, center)) {
                patchedCount++;
            }
            healedCount++;
        }

        try { gapLayer.locked = gapLayerWasLocked; } catch (e) { }
        try { gapLayer.visible = gapLayerWasVisible; } catch (e) { }

        if (restoredSegments.length > 0) {
            try {
                doc.selection = null;
                for (var rs = 0; rs < restoredSegments.length; rs++) {
                    try { restoredSegments[rs].selected = true; } catch (e) { }
                }
                MDUX_mergePathsAtEndpoints();
            } catch (e) { }
        }

        if (markers.length < 1) return "No gap markers near selection";
        if (healedCount < 1 && skippedMissing > 0) return "No gaps healed (missing deleted segments)";

        return "Healed " + healedCount + " gap(s), restored " + restoredCount + ", removed " + markerRemoved +
            " marker(s), patched " + patchedCount + ", removed " + ignoreRemoved + " ignore anchor(s)" +
            (skippedMissing > 0 ? ", skipped " + skippedMissing : "");
    } catch (e) {
        return "Error: " + e;
    }
}

function MDUX_recreateGapsInSelection() {
    try {
        if (app.documents.length === 0) return "No document open";
        var doc = app.activeDocument;
        var selectedPaths = MDUX_gapTools_collectSelectedPaths(doc);
        if (!selectedPaths || selectedPaths.length < 1) return "Select ductwork paths to recreate gaps";

        var bounds = MDUX_gapTools_getSelectionBounds(selectedPaths);
        var gapLayer = MDUX_gapTools_getGapLayer(doc, true);
        if (!gapLayer) return "Unable to access Gap Definitions layer";

        var gapLayerWasLocked = gapLayer.locked;
        var gapLayerWasVisible = gapLayer.visible;
        try { if (gapLayerWasLocked) gapLayer.locked = false; } catch (e) { }
        try { if (!gapLayerWasVisible) gapLayer.visible = true; } catch (e) { }

        var ignoreLayer = MDUX_gapTools_getIgnoreLayer(doc, true);
        var ignoreWasLocked = ignoreLayer ? ignoreLayer.locked : false;
        var ignoreWasVisible = ignoreLayer ? ignoreLayer.visible : true;
        try { if (ignoreLayer && ignoreWasLocked) ignoreLayer.locked = false; } catch (e) { }
        try { if (ignoreLayer && !ignoreWasVisible) ignoreLayer.visible = true; } catch (e) { }

        var deletedLayer = null;
        try { deletedLayer = doc.layers.getByName("Deleted Segments"); } catch (e) { deletedLayer = null; }

        var targets = [];
        for (var i = 0; i < gapLayer.pathItems.length; i++) {
            var item = gapLayer.pathItems[i];
            if (!item || !item.pathPoints || item.pathPoints.length < 1) continue;
            var note = item.note || "";
            var type = null;
            if (note.indexOf("MDUX_GAP:") === 0) type = "gap";
            else if (note.indexOf("MDUX_PATCH") === 0) type = "patch";
            else type = "anchor";

            var info = (type === "gap") ? MDUX_gapTools_getMarkerInfo(item) : null;
            var center = null;
            if (info && info.center) {
                center = info.center;
            } else {
                try { center = [item.pathPoints[0].anchor[0], item.pathPoints[0].anchor[1]]; } catch (e) { center = null; }
            }
            if (!center) continue;
            if (!MDUX_gapTools_pointInBounds(center, bounds, 25)) continue;
            targets.push({ item: item, type: type, center: center, info: info });
        }

        var activePaths = selectedPaths.slice();
        var recreated = 0;
        var skipped = 0;

        for (var t = 0; t < targets.length; t++) {
            var target = targets[t];
            var center = target.center;
            var info = target.info || {};
            var sourceLayer = info.sourceLayer || null;

            var near = MDUX_gapTools_findNearestSegment(activePaths, center, sourceLayer, 25);
            if (!near) { skipped++; continue; }

            var segLen = near.segLen;
            if (!segLen || segLen < 0.01) { skipped++; continue; }

            var gapSize = info.gapSize;
            var isAutoSized = info.isAutoSized;
            if (!gapSize || gapSize <= 0) {
                var strokeWidth = null;
                try { strokeWidth = near.path.strokeWidth; } catch (e) { strokeWidth = null; }
                gapSize = MDUX_gapTools_calculateAutoGapSize(strokeWidth);
                isAutoSized = true;
            }

            var dir = info.dir;
            if (!dir || dir.length < 2) {
                var dx = near.segEnd[0] - near.segStart[0];
                var dy = near.segEnd[1] - near.segStart[1];
                var len = Math.sqrt(dx * dx + dy * dy);
                if (len < 0.01) { skipped++; continue; }
                dir = [dx / len, dy / len];
            }

            var ratio = gapSize / segLen;
            if (ratio >= 0.49) { skipped++; continue; }
            if (near.t <= ratio || near.t >= (1 - ratio)) { skipped++; continue; }

            var centerOnSeg = near.proj;
            var cutBefore = [centerOnSeg[0] - gapSize * dir[0], centerOnSeg[1] - gapSize * dir[1]];
            var cutAfter = [centerOnSeg[0] + gapSize * dir[0], centerOnSeg[1] + gapSize * dir[1]];

            if (deletedLayer) {
                var existingSeg = MDUX_gapTools_findDeletedSegmentNear(deletedLayer, centerOnSeg, 15);
                if (existingSeg) { skipped++; continue; }
            }

            MDUX_gapTools_saveDeletedSegment(doc, cutBefore, cutAfter, sourceLayer || (near.path.layer ? near.path.layer.name : ""));

            if (target.type === "gap") {
                MDUX_gapTools_createGapMarker(doc, gapLayer, centerOnSeg, gapSize, sourceLayer, dir, isAutoSized, target.item);
            } else {
                if (target.type === "patch") {
                    try { target.item.locked = false; } catch (e) { }
                    try { target.item.remove(); } catch (e) { }
                }
                MDUX_gapTools_createGapMarker(doc, gapLayer, centerOnSeg, gapSize, sourceLayer, dir, isAutoSized, null);
            }

            if (ignoreLayer) {
                MDUX_gapTools_createIgnoreAnchor(ignoreLayer, cutBefore);
                MDUX_gapTools_createIgnoreAnchor(ignoreLayer, cutAfter);
            }

            var parentLayer = MDUX_gapTools_getParentLayer(near.path);
            var split = MDUX_gapTools_splitPathAtGap(near.path, near.segIdx, cutBefore, cutAfter, parentLayer);
            if (split && split.first && split.second) {
                for (var ap = activePaths.length - 1; ap >= 0; ap--) {
                    if (activePaths[ap] === near.path) {
                        activePaths.splice(ap, 1);
                    }
                }
                try { if (split.first.pathPoints && split.first.pathPoints.length >= 2) activePaths.push(split.first); } catch (e) { }
                try { if (split.second.pathPoints && split.second.pathPoints.length >= 2) activePaths.push(split.second); } catch (e) { }
                recreated++;
            } else {
                skipped++;
            }
        }

        try { gapLayer.locked = gapLayerWasLocked; } catch (e) { }
        try { gapLayer.visible = gapLayerWasVisible; } catch (e) { }

        try {
            if (ignoreLayer) {
                ignoreLayer.locked = ignoreWasLocked;
                ignoreLayer.visible = ignoreWasVisible;
            }
        } catch (e) { }

        if (targets.length < 1) return "No gap markers near selection";
        return "Recreated " + recreated + " gap(s)" + (skipped > 0 ? ", skipped " + skipped : "");
    } catch (e) {
        return "Error: " + e;
    }
}
