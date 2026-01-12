// Gap tools for manual heal/recreate workflows
var MDUX_gapTools_traceLines = null;
var MDUX_gapTools_traceTag = null;
var MDUX_gapTools_buildTag = "2026-01-12-1156";

function MDUX_gapTools_isTraceEnabled() {
    try {
        if ($.global && $.global.MDUX_DEBUG && $.global.MDUX_DEBUG.GAP_TRACE === true) {
            return true;
        }
    } catch (e) { }
    return false;
}

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

function MDUX_gapTools_getParentCompound(item) {
    var cur = item;
    while (cur) {
        try {
            if (cur.typename === "CompoundPathItem") return cur;
            cur = cur.parent;
        } catch (e) {
            cur = null;
        }
    }
    return null;
}

function MDUX_gapTools_isValidPath(item) {
    if (!item) return false;
    try {
        if (item.isValid === false) return false;
    } catch (e) { return false; }
    try {
        if (item.typename !== "PathItem") return false;
    } catch (e2) { return false; }
    return true;
}

function MDUX_gapTools_isValidLayer(layer) {
    if (!layer) return false;
    try {
        if (layer.isValid === false) return false;
    } catch (e) { return false; }
    try {
        if (layer.typename !== "Layer") return false;
    } catch (e2) { return false; }
    return true;
}

function MDUX_gapTools_safeLayerName(layer) {
    try { return (layer && layer.name) ? layer.name : "null"; } catch (e) { return "invalid"; }
}

function MDUX_gapTools_filterValidPaths(paths) {
    var out = [];
    if (!paths) return out;
    for (var i = 0; i < paths.length; i++) {
        var p = paths[i];
        if (!p) continue;
        if (!MDUX_gapTools_isValidPath(p)) continue;
        out.push(p);
    }
    return out;
}

function MDUX_gapTools_debug(message) {
    try {
        if (typeof MDUX_debugLog === "function") {
            var tag = "";
            try {
                if (MDUX_gapTools_traceTag) tag = MDUX_gapTools_traceTag + " ";
                else if ($.global && $.global.MDUX_GAP_TRACE_TAG) {
                    tag = $.global.MDUX_GAP_TRACE_TAG + " ";
                }
            } catch (e) { }
            MDUX_debugLog("[GAP-TOOLS] " + tag + message);
        }
    } catch (e) { }
}

function MDUX_gapTools_setTraceTag(tag) {
    MDUX_gapTools_traceTag = tag || "";
    try { $.global.MDUX_GAP_TRACE_TAG = MDUX_gapTools_traceTag; } catch (e) { }
}

function MDUX_gapTools_clearTraceTag() {
    MDUX_gapTools_traceTag = null;
    try { $.global.MDUX_GAP_TRACE_TAG = null; } catch (e) { }
}

function MDUX_gapTools_setTraceBuffer(buffer) {
    MDUX_gapTools_traceLines = buffer || null;
    try { $.global.MDUX_GAP_TRACE_LINES = MDUX_gapTools_traceLines; } catch (e) { }
}

function MDUX_gapTools_clearTraceBuffer() {
    MDUX_gapTools_traceLines = null;
    try { $.global.MDUX_GAP_TRACE_LINES = null; } catch (e) { }
}

function MDUX_gapTools_trace(message) {
    if (!MDUX_gapTools_isTraceEnabled()) return;
    try {
        var lines = MDUX_gapTools_traceLines;
        if (!lines) {
            try {
                if ($.global && $.global.MDUX_GAP_TRACE_LINES && $.global.MDUX_GAP_TRACE_LINES.push) {
                    lines = $.global.MDUX_GAP_TRACE_LINES;
                }
            } catch (e) { }
        }
        if (lines && lines.push) lines.push(message);
    } catch (e) { }
    MDUX_gapTools_debug(message);
}

function MDUX_gapTools_traceError(prefix, err) {
    try {
        if (typeof MDUX_debugLog === "function") {
            MDUX_debugLog("[GAP-TOOLS] " + prefix + ": " + err);
        }
        if (MDUX_gapTools_isTraceEnabled()) {
            MDUX_gapTools_trace(prefix + ": " + err);
        }
    } catch (e) { }
}

function MDUX_gapTools_formatPoint(pt) {
    if (!pt) return "null";
    return pt[0].toFixed(2) + "," + pt[1].toFixed(2);
}

function MDUX_gapTools_formatSegment(path) {
    try {
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) return "segment?(invalid)";
        var a = pts[0].anchor;
        var b = pts[pts.length - 1].anchor;
        return "seg[" + MDUX_gapTools_formatPoint(a) + " -> " + MDUX_gapTools_formatPoint(b) + "]";
    } catch (e) {
        return "segment?(error)";
    }
}

function MDUX_gapTools_writeTrace(tag, traceLines) {
    if (!MDUX_gapTools_isTraceEnabled()) return "";
    if (!traceLines) traceLines = [];
    try {
        var folderPath = "C:/Users/Chris/AppData/Roaming/Adobe/CEP/extensions/Magic-Ductwork-Panel/Debug";
        var debugFolder = new Folder(folderPath);
        if (!debugFolder.exists) {
            debugFolder.create();
        }
        var cleanTag = String(tag || "gap").replace(/[^A-Za-z0-9_-]/g, "_");
        var timestamp = new Date();
        var dateStr = timestamp.getFullYear() + "-" +
                     ("0" + (timestamp.getMonth() + 1)).slice(-2) + "-" +
                     ("0" + timestamp.getDate()).slice(-2) + "_" +
                     ("0" + timestamp.getHours()).slice(-2) + "-" +
                     ("0" + timestamp.getMinutes()).slice(-2) + "-" +
                     ("0" + timestamp.getSeconds()).slice(-2);
        var filePath = folderPath + "/gap-trace-" + cleanTag + "-" + dateStr + ".log";
        var outFile = new File(filePath);
        if (outFile.open("w")) {
            outFile.encoding = "UTF-8";
            var payload = traceLines.length > 0 ? traceLines.join("\n") : "(no trace lines captured)";
            outFile.write("=== GAP TRACE " + (tag || "") + " ===\n" + payload);
            outFile.close();
            return filePath;
        }
    } catch (e) { }
    return "";
}

function MDUX_gapTools_finalizeTrace(tag, traceLines) {
    var traceFile = MDUX_gapTools_writeTrace(tag, traceLines);
    MDUX_gapTools_clearTraceTag();
    MDUX_gapTools_clearTraceBuffer();
    return traceFile;
}

function MDUX_gapTools_mergeWithinCompound(compound, tol) {
    if (!compound || !compound.pathItems) return 0;
    var paths = [];
    try {
        for (var i = 0; i < compound.pathItems.length; i++) {
            var p = compound.pathItems[i];
            if (!p.closed && p.pathPoints && p.pathPoints.length >= 2) {
                paths.push(p);
            }
        }
    } catch (e) { }
    if (paths.length < 2) return 0;

    var removed = [];
    function wasRemoved(path) {
        for (var r = 0; r < removed.length; r++) {
            if (removed[r] === path) return true;
        }
        return false;
    }

    function getEndpoints(path) {
        try {
            var pts = path.pathPoints;
            if (!pts || pts.length < 2) return null;
            return {
                start: [pts[0].anchor[0], pts[0].anchor[1]],
                end: [pts[pts.length - 1].anchor[0], pts[pts.length - 1].anchor[1]]
            };
        } catch (e) { return null; }
    }

    function mergePaths(pathA, pathB, matchType) {
        var ptsA = pathA.pathPoints;
        var ptsB = pathB.pathPoints;
        var allPoints = [];

        if (matchType === "A_end_B_start") {
            for (var i = 0; i < ptsA.length - 1; i++) allPoints.push([ptsA[i].anchor[0], ptsA[i].anchor[1]]);
            for (var i2 = 1; i2 < ptsB.length; i2++) allPoints.push([ptsB[i2].anchor[0], ptsB[i2].anchor[1]]);
        } else if (matchType === "A_end_B_end") {
            for (var j = 0; j < ptsA.length - 1; j++) allPoints.push([ptsA[j].anchor[0], ptsA[j].anchor[1]]);
            for (var j2 = ptsB.length - 2; j2 >= 0; j2--) allPoints.push([ptsB[j2].anchor[0], ptsB[j2].anchor[1]]);
        } else if (matchType === "A_start_B_end") {
            for (var k = 0; k < ptsB.length - 1; k++) allPoints.push([ptsB[k].anchor[0], ptsB[k].anchor[1]]);
            for (var k2 = 1; k2 < ptsA.length; k2++) allPoints.push([ptsA[k2].anchor[0], ptsA[k2].anchor[1]]);
        } else if (matchType === "A_start_B_start") {
            for (var m = ptsB.length - 1; m >= 1; m--) allPoints.push([ptsB[m].anchor[0], ptsB[m].anchor[1]]);
            for (var m2 = 1; m2 < ptsA.length; m2++) allPoints.push([ptsA[m2].anchor[0], ptsA[m2].anchor[1]]);
        }

        pathA.setEntirePath(allPoints);
        return pathA;
    }

    var tolerance = (typeof tol === "number") ? tol : 1.0;
    var mergedCount = 0;
    var keepMerging = true;
    var iterations = 0;
    var maxIterations = paths.length * 3;

    while (keepMerging && iterations < maxIterations) {
        keepMerging = false;
        iterations++;

        for (var a = 0; a < paths.length; a++) {
            var pathA = paths[a];
            if (!pathA || wasRemoved(pathA)) continue;
            var endpointsA = getEndpoints(pathA);
            if (!endpointsA) continue;

            for (var b = 0; b < paths.length; b++) {
                var pathB = paths[b];
                if (!pathB || pathB === pathA || wasRemoved(pathB)) continue;
                var endpointsB = getEndpoints(pathB);
                if (!endpointsB) continue;

                var d1 = MDUX_gapTools_pointDistance(endpointsA.end, endpointsB.start);
                var d2 = MDUX_gapTools_pointDistance(endpointsA.end, endpointsB.end);
                var d3 = MDUX_gapTools_pointDistance(endpointsA.start, endpointsB.end);
                var d4 = MDUX_gapTools_pointDistance(endpointsA.start, endpointsB.start);

                var minDist = Math.min(d1, d2, d3, d4);
                if (minDist > tolerance) continue;

                var matchType = "A_end_B_start";
                if (minDist === d1) matchType = "A_end_B_start";
                else if (minDist === d2) matchType = "A_end_B_end";
                else if (minDist === d3) matchType = "A_start_B_end";
                else if (minDist === d4) matchType = "A_start_B_start";

                try {
                    mergePaths(pathA, pathB, matchType);
                    pathB.remove();
                    removed.push(pathB);
                    mergedCount++;
                    keepMerging = true;
                } catch (mergeErr) { }

                break;
            }
            if (keepMerging) break;
        }
    }

    return mergedCount;
}

function MDUX_gapTools_getEndpointDir(path, atStart) {
    try {
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) return [1, 0];
        var a = null;
        var b = null;
        if (atStart) {
            a = pts[0].anchor;
            b = pts[1].anchor;
        } else {
            a = pts[pts.length - 2].anchor;
            b = pts[pts.length - 1].anchor;
        }
        var dx = b[0] - a[0];
        var dy = b[1] - a[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len < 0.001) return [1, 0];
        return [dx / len, dy / len];
    } catch (e) { return [1, 0]; }
}

function MDUX_gapTools_collectOpenPathsFromCompound(compound) {
    var paths = [];
    if (!compound || !compound.pathItems) return paths;
    try {
        for (var i = 0; i < compound.pathItems.length; i++) {
            var p = compound.pathItems[i];
            if (!MDUX_gapTools_isValidPath(p)) continue;
            if (!p.closed && p.pathPoints && p.pathPoints.length >= 2) {
                paths.push(p);
            }
        }
    } catch (e) { }
    return paths;
}

function MDUX_gapTools_mergePathWithNeighbors(path, candidates, tol, dotThreshold) {
    if (!path || !candidates || candidates.length < 1) return 0;
    if (!MDUX_gapTools_isValidPath(path)) {
        MDUX_gapTools_trace("MERGE abort invalid path");
        return 0;
    }
    var tolerance = (typeof tol === "number") ? tol : 3.0;
    var minDot = (typeof dotThreshold === "number") ? dotThreshold : 0.985;
    var liveCandidates = MDUX_gapTools_filterValidPaths(candidates);

    MDUX_gapTools_trace("MERGE start " + MDUX_gapTools_formatSegment(path) + " candidates=" + liveCandidates.length);

    function getEndpoints(p) {
        try {
            var pts = p.pathPoints;
            if (!pts || pts.length < 2) return null;
            return {
                start: [pts[0].anchor[0], pts[0].anchor[1]],
                end: [pts[pts.length - 1].anchor[0], pts[pts.length - 1].anchor[1]]
            };
        } catch (e) { return null; }
    }

    function copyStyle(src, dst) {
        if (!src || !dst) return;
        try { dst.filled = src.filled; } catch (e) { }
        try { dst.stroked = src.stroked; } catch (e) { }
        try { dst.strokeWidth = src.strokeWidth; } catch (e) { }
        try { dst.strokeColor = src.strokeColor; } catch (e) { }
        try { if (src.filled) dst.fillColor = src.fillColor; } catch (e) { }
    }

    function getPointsCopy(path) {
        try {
            var pts = path.pathPoints;
            if (!pts || pts.length < 2) return null;
            var out = [];
            for (var i = 0; i < pts.length; i++) {
                out.push([pts[i].anchor[0], pts[i].anchor[1]]);
            }
            return out;
        } catch (e) { return null; }
    }

    function buildMergedPoints(pointsA, pointsB, matchType) {
        var allPoints = [];
        if (!pointsA || !pointsB) return null;
        if (matchType === "A_end_B_start") {
            for (var i = 0; i < pointsA.length - 1; i++) allPoints.push(pointsA[i]);
            for (var i2 = 0; i2 < pointsB.length; i2++) allPoints.push(pointsB[i2]);
        } else if (matchType === "A_end_B_end") {
            for (var j = 0; j < pointsA.length - 1; j++) allPoints.push(pointsA[j]);
            for (var j2 = pointsB.length - 1; j2 >= 0; j2--) allPoints.push(pointsB[j2]);
        } else if (matchType === "A_start_B_end") {
            for (var k = 0; k < pointsB.length - 1; k++) allPoints.push(pointsB[k]);
            for (var k2 = 0; k2 < pointsA.length; k2++) allPoints.push(pointsA[k2]);
        } else if (matchType === "A_start_B_start") {
            for (var m = pointsB.length - 1; m >= 1; m--) allPoints.push(pointsB[m]);
            for (var m2 = 0; m2 < pointsA.length; m2++) allPoints.push(pointsA[m2]);
        }
        return allPoints;
    }

    function createMergedPathLike(src, points) {
        try {
            if (!MDUX_gapTools_isValidPath(src)) return null;
            var layer = null;
            try { layer = MDUX_gapTools_getParentLayer(src); } catch (e) { layer = null; }
            try { if (!layer && src.layer) layer = src.layer; } catch (e) { }
            if (!layer || !MDUX_gapTools_isValidLayer(layer)) {
                try { layer = app.activeDocument.activeLayer; } catch (e) { layer = null; }
            }
            if (!layer || !MDUX_gapTools_isValidLayer(layer)) return null;
            var newPath = null;
            try { newPath = layer.pathItems.add(); } catch (e) { newPath = null; }
            if (!newPath) return null;
            try { newPath.setEntirePath(points); } catch (e) { }
            copyStyle(src, newPath);
            var compound = null;
            try { compound = MDUX_gapTools_getParentCompound(src); } catch (e) { compound = null; }
            if (compound) {
                try { newPath.move(compound, ElementPlacement.PLACEATEND); } catch (e) { }
            }
            return newPath;
        } catch (e) {
            return null;
        }
    }

    function mergePaths(pathA, pathB, matchType) {
        try {
            if (!MDUX_gapTools_isValidPath(pathA) || !MDUX_gapTools_isValidPath(pathB)) return null;
            var pointsA = getPointsCopy(pathA);
            var pointsB = getPointsCopy(pathB);
            if (!pointsA || !pointsB) return null;
            var mergedPoints = buildMergedPoints(pointsA, pointsB, matchType);
            if (!mergedPoints || mergedPoints.length < 2) return null;
            try {
                pathA.setEntirePath(mergedPoints);
                return pathA;
            } catch (e) {
                var newPath = null;
                try { newPath = createMergedPathLike(pathA, mergedPoints); } catch (e2) { newPath = null; }
                if (newPath) {
                    try { pathA.remove(); } catch (e3) { }
                    return newPath;
                }
            }
        } catch (eAll) {
            return null;
        }
        return null;
    }

    var mergedCount = 0;
    var keepMerging = true;
    var iterations = 0;
    var maxIterations = liveCandidates.length * 3;

    while (keepMerging && iterations < maxIterations) {
        keepMerging = false;
        iterations++;

        if (!MDUX_gapTools_isValidPath(path)) break;
        var endpointsA = getEndpoints(path);
        if (!endpointsA) break;

        var endpointsInfo = [
            { pt: endpointsA.start, isStart: true, dir: MDUX_gapTools_getEndpointDir(path, true) },
            { pt: endpointsA.end, isStart: false, dir: MDUX_gapTools_getEndpointDir(path, false) }
        ];

        for (var eIdx = 0; eIdx < endpointsInfo.length; eIdx++) {
            var endpointInfo = endpointsInfo[eIdx];
            var matches = [];
            MDUX_gapTools_trace("MERGE endpoint " + (endpointInfo.isStart ? "start" : "end") +
                " pt=" + MDUX_gapTools_formatPoint(endpointInfo.pt));

            for (var c = 0; c < liveCandidates.length; c++) {
                var cand = liveCandidates[c];
                if (!cand) continue;
                if (!MDUX_gapTools_isValidPath(cand)) {
                    liveCandidates[c] = null;
                    continue;
                }
                if (cand === path) continue;
                var endpointsB = getEndpoints(cand);
                if (!endpointsB) continue;

                var dStart = MDUX_gapTools_pointDistance(endpointInfo.pt, endpointsB.start);
                var dEnd = MDUX_gapTools_pointDistance(endpointInfo.pt, endpointsB.end);
                if (dStart > tolerance && dEnd > tolerance) continue;

            function pushMatch(matchAtStart, distVal, candIndex) {
                var matchType = "A_end_B_start";
                if (endpointInfo.isStart && matchAtStart) matchType = "A_start_B_start";
                else if (endpointInfo.isStart && !matchAtStart) matchType = "A_start_B_end";
                else if (!endpointInfo.isStart && matchAtStart) matchType = "A_end_B_start";
                else if (!endpointInfo.isStart && !matchAtStart) matchType = "A_end_B_end";

                    var candDir = MDUX_gapTools_getEndpointDir(cand, matchAtStart);
                    var signedDot = endpointInfo.dir[0] * candDir[0] + endpointInfo.dir[1] * candDir[1];
                    var absDot = Math.abs(signedDot);
                    var expectedSign = (matchAtStart !== endpointInfo.isStart) ? 1 : -1;
                    var score = expectedSign * signedDot;
                matches.push({
                    path: cand,
                    matchType: matchType,
                    dot: absDot,
                    signedDot: signedDot,
                    score: score,
                    dist: distVal,
                    candIndex: candIndex
                });
            }

                if (dStart <= tolerance && dEnd <= tolerance) {
                    pushMatch(true, dStart, c);
                    pushMatch(false, dEnd, c);
                } else if (dStart <= tolerance) {
                    pushMatch(true, dStart, c);
                } else if (dEnd <= tolerance) {
                    pushMatch(false, dEnd, c);
                }
            }

            if (matches.length < 1) {
                MDUX_gapTools_trace("MERGE no matches within tol=" + tolerance);
                continue;
            }

            matches.sort(function(a, b) {
                if (Math.abs(b.score - a.score) > 1e-6) return b.score - a.score;
                return a.dist - b.dist;
            });

            var best = matches[0];
            var second = matches.length > 1 ? matches[1] : null;
            if (best.dot < minDot || best.score < minDot) {
                MDUX_gapTools_trace("MERGE skip best.dot=" + best.dot.toFixed(3) +
                    " score=" + best.score.toFixed(3) + " < " + minDot.toFixed(3));
                continue;
            }
            if (second && Math.abs(best.score - second.score) < 0.01 && Math.abs(best.dist - second.dist) < 0.1) {
                MDUX_gapTools_trace("MERGE tie best.score=" + best.score.toFixed(3) +
                    " second.score=" + second.score.toFixed(3) + " dist=" + best.dist.toFixed(2));
            }

            var mergeErrStage = "init";
            try {
                mergeErrStage = "validate-path";
                if (!MDUX_gapTools_isValidPath(path)) {
                    MDUX_gapTools_trace("MERGE abort path invalid before merge");
                    return mergedCount;
                }
                mergeErrStage = "validate-candidate";
                if (!MDUX_gapTools_isValidPath(best.path)) {
                    MDUX_gapTools_trace("MERGE skip invalid candidate");
                    if (typeof best.candIndex === "number" && best.candIndex >= 0 && best.candIndex < liveCandidates.length) {
                        liveCandidates[best.candIndex] = null;
                    }
                    continue;
                }
                mergeErrStage = "format-seg";
                var candLabel = "segment?";
                try { candLabel = MDUX_gapTools_formatSegment(best.path); } catch (eLbl) { candLabel = "segment?(invalid)"; }
                MDUX_gapTools_trace("MERGE chosen match=" + best.matchType +
                    " score=" + best.score.toFixed(3) + " dot=" + best.dot.toFixed(3) +
                    " signed=" + best.signedDot.toFixed(3) + " dist=" + best.dist.toFixed(2) +
                    " cand=" + candLabel);
                mergeErrStage = "mergePaths";
                var mergedPath = mergePaths(path, best.path, best.matchType);
                if (!mergedPath) {
                    MDUX_gapTools_trace("MERGE failed to apply merge");
                    continue;
                }
                path = mergedPath;
                mergeErrStage = "cleanup";
                var candIndex = (typeof best.candIndex === "number") ? best.candIndex : -1;
                if (candIndex >= 0 && candIndex < liveCandidates.length) {
                    liveCandidates[candIndex] = null;
                }
                mergeErrStage = "remove-candidate";
                try { best.path.remove(); } catch (eRemove) {
                    MDUX_gapTools_traceError("MERGE remove error", eRemove);
                    keepMerging = false;
                    break;
                }
                mergedCount++;
                keepMerging = true;
                MDUX_gapTools_trace("MERGE success mergedCount=" + mergedCount);
                var simplified = MDUX_gapTools_simplifyCollinear(path, 0.5, 0.999);
                if (simplified > 0) {
                    MDUX_gapTools_trace("MERGE simplified=" + simplified);
                }
            } catch (mergeErr) {
                MDUX_gapTools_traceError("MERGE error @" + mergeErrStage, mergeErr);
            }

            break;
        }
    }

    return mergedCount;
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

function MDUX_gapTools_isStraightPoint(point, tol) {
    if (!point) return false;
    var t = (typeof tol === "number") ? tol : 0.01;
    try {
        var a = point.anchor;
        var l = point.leftDirection;
        var r = point.rightDirection;
        if (!a || !l || !r) return false;
        if (MDUX_gapTools_pointDistance(a, l) > t) return false;
        if (MDUX_gapTools_pointDistance(a, r) > t) return false;
        return true;
    } catch (e) {
        return false;
    }
}

function MDUX_gapTools_simplifyCollinear(path, distTol, dotThreshold) {
    if (!MDUX_gapTools_isValidPath(path)) return 0;
    var pts = null;
    try { pts = path.pathPoints; } catch (e) { pts = null; }
    if (!pts || pts.length < 3) return 0;
    var distLimit = (typeof distTol === "number") ? distTol : 0.5;
    var minDot = (typeof dotThreshold === "number") ? dotThreshold : 0.999;
    var handleTol = 0.01;
    var removed = 0;

    for (var i = pts.length - 2; i >= 1; i--) {
        var prev = null, curr = null, next = null;
        try {
            prev = pts[i - 1];
            curr = pts[i];
            next = pts[i + 1];
        } catch (e) { continue; }
        if (!prev || !curr || !next) continue;
        if (!MDUX_gapTools_isStraightPoint(curr, handleTol)) continue;
        if (!MDUX_gapTools_isStraightPoint(prev, handleTol)) continue;
        if (!MDUX_gapTools_isStraightPoint(next, handleTol)) continue;

        var a = prev.anchor;
        var b = curr.anchor;
        var c = next.anchor;
        if (!a || !b || !c) continue;
        var v1x = b[0] - a[0];
        var v1y = b[1] - a[1];
        var v2x = c[0] - b[0];
        var v2y = c[1] - b[1];
        var len1 = Math.sqrt(v1x * v1x + v1y * v1y);
        var len2 = Math.sqrt(v2x * v2x + v2y * v2y);
        if (len1 < 0.01 || len2 < 0.01) continue;
        var dot = (v1x * v2x + v1y * v2y) / (len1 * len2);
        if (dot < minDot) continue;

        var proj = MDUX_gapTools_closestPointOnSegment(b, a, c);
        if (!proj || proj.dist > distLimit) continue;

        try { curr.remove(); removed++; } catch (e) { }
    }

    return removed;
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

function MDUX_gapTools_getPathLength(path) {
    if (!path) return 0;
    var length = 0;
    try {
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) return 0;
        for (var i = 0; i < pts.length - 1; i++) {
            var a = pts[i].anchor;
            var b = pts[i + 1].anchor;
            var dx = b[0] - a[0];
            var dy = b[1] - a[1];
            length += Math.sqrt(dx * dx + dy * dy);
        }
    } catch (e) { }
    return length;
}

function MDUX_gapTools_getSegmentIntersection(ax1, ay1, ax2, ay2, bx1, by1, bx2, by2) {
    var dax = ax2 - ax1, day = ay2 - ay1;
    var dbx = bx2 - bx1, dby = by2 - by1;
    var denom = dax * dby - day * dbx;
    if (Math.abs(denom) < 1e-10) return null;
    var dx = ax1 - bx1, dy = ay1 - by1;
    var t = (dbx * dy - dby * dx) / denom;
    var u = (dax * dy - day * dx) / denom;
    if (t < 0 || t > 1 || u < 0 || u > 1) return null;
    return { pt: [ax1 + t * dax, ay1 + t * day], t1: t, t2: u };
}

function MDUX_gapTools_collectIntersectionCenters(paths, endpointTol) {
    var centers = [];
    var seen = {};
    var tol = (typeof endpointTol === "number") ? endpointTol : 1;
    if (!paths || paths.length < 2) return centers;

    for (var i = 0; i < paths.length; i++) {
        var pathA = paths[i];
        var ptsA = null;
        try { ptsA = pathA.pathPoints; } catch (e) { ptsA = null; }
        if (!ptsA || ptsA.length < 2) continue;

        for (var j = i + 1; j < paths.length; j++) {
            var pathB = paths[j];
            var ptsB = null;
            try { ptsB = pathB.pathPoints; } catch (e2) { ptsB = null; }
            if (!ptsB || ptsB.length < 2) continue;

            try {
                if (pathA.layer && pathB.layer && pathA.layer.name !== pathB.layer.name) continue;
            } catch (eLayer) { }

            for (var ai = 0; ai < ptsA.length - 1; ai++) {
                var a0 = ptsA[ai].anchor;
                var a1 = ptsA[ai + 1].anchor;
                for (var bi = 0; bi < ptsB.length - 1; bi++) {
                    var b0 = ptsB[bi].anchor;
                    var b1 = ptsB[bi + 1].anchor;
                    var inter = MDUX_gapTools_getSegmentIntersection(
                        a0[0], a0[1], a1[0], a1[1],
                        b0[0], b0[1], b1[0], b1[1]
                    );
                    if (!inter) continue;
                    var aDx = a1[0] - a0[0];
                    var aDy = a1[1] - a0[1];
                    var bDx = b1[0] - b0[0];
                    var bDy = b1[1] - b0[1];
                    var segLenA = Math.sqrt(aDx * aDx + aDy * aDy);
                    var segLenB = Math.sqrt(bDx * bDx + bDy * bDy);
                    if (segLenA < 0.01 || segLenB < 0.01) continue;

                    var distAStart = inter.t1 * segLenA;
                    var distAEnd = (1 - inter.t1) * segLenA;
                    var distBStart = inter.t2 * segLenB;
                    var distBEnd = (1 - inter.t2) * segLenB;

                    var aIsPathEndpoint = ((distAStart <= tol) && ai === 0) ||
                        ((distAEnd <= tol) && (ai + 1 === ptsA.length - 1));
                    var bIsPathEndpoint = ((distBStart <= tol) && bi === 0) ||
                        ((distBEnd <= tol) && (bi + 1 === ptsB.length - 1));

                    if (aIsPathEndpoint || bIsPathEndpoint) continue;

                    var key = Math.round(inter.pt[0] * 10) / 10 + "_" + Math.round(inter.pt[1] * 10) / 10;
                    if (seen[key]) continue;
                    seen[key] = true;
                    centers.push(inter.pt);
                }
            }
        }
    }
    return centers;
}

function MDUX_gapTools_findSegmentsNearPoint(paths, point, maxDist) {
    var hits = [];
    if (!paths || !point) return hits;
    var distLimit = (typeof maxDist === "number") ? maxDist : 1;

    for (var i = 0; i < paths.length; i++) {
        var path = paths[i];
        var pts = null;
        try { pts = path.pathPoints; } catch (e) { pts = null; }
        if (!pts || pts.length < 2) continue;

        for (var s = 0; s < pts.length - 1; s++) {
            var a = pts[s].anchor;
            var b = pts[s + 1].anchor;
            var res = MDUX_gapTools_closestPointOnSegment(point, a, b);
            if (res.segLen < 0.01) continue;
            if (res.dist <= distLimit) {
                var dx = b[0] - a[0];
                var dy = b[1] - a[1];
                var len = res.segLen;
                var dir = (len > 0) ? [dx / len, dy / len] : [1, 0];
                hits.push({
                    path: path,
                    segIdx: s,
                    proj: res.pt,
                    dist: res.dist,
                    segStart: [a[0], a[1]],
                    segEnd: [b[0], b[1]],
                    segLen: res.segLen,
                    t: res.t,
                    dir: dir
                });
            }
        }
    }
    return hits;
}

function MDUX_gapTools_findIntersectingSegmentsAtPoint(paths, point, maxDist) {
    var hits = MDUX_gapTools_findSegmentsNearPoint(paths, point, maxDist);
    if (hits.length < 1) return [];

    var bestByPath = [];
    for (var i = 0; i < hits.length; i++) {
        var hit = hits[i];
        var found = false;
        for (var j = 0; j < bestByPath.length; j++) {
            if (bestByPath[j].path === hit.path) {
                found = true;
                if (hit.dist < bestByPath[j].dist) {
                    bestByPath[j] = hit;
                }
                break;
            }
        }
        if (!found) bestByPath.push(hit);
    }

    bestByPath.sort(function(a, b) { return a.dist - b.dist; });
    return bestByPath;
}

function MDUX_gapTools_chooseBestPairByDirection(hits) {
    if (!hits || hits.length < 2) return null;
    var best = null;
    var bestDot = 1e9;
    var bestDist = 1e9;
    for (var i = 0; i < hits.length - 1; i++) {
        for (var j = i + 1; j < hits.length; j++) {
            var dot = Math.abs(hits[i].dir[0] * hits[j].dir[0] + hits[i].dir[1] * hits[j].dir[1]);
            var avgDist = (hits[i].dist + hits[j].dist) / 2;
            if (dot < bestDot - 1e-6 || (Math.abs(dot - bestDot) < 1e-6 && avgDist < bestDist)) {
                best = { a: hits[i], b: hits[j], dot: dot };
                bestDot = dot;
                bestDist = avgDist;
            }
        }
    }
    return best;
}

function MDUX_gapTools_chooseFlipHits(hits, segDir) {
    if (!hits || hits.length < 1 || !segDir) return null;
    var restoreHit = null;
    var restoreDot = -1;
    for (var i = 0; i < hits.length; i++) {
        var dot = Math.abs(segDir[0] * hits[i].dir[0] + segDir[1] * hits[i].dir[1]);
        if (dot > restoreDot) {
            restoreDot = dot;
            restoreHit = hits[i];
        }
    }
    if (!restoreHit) return null;
    var gapHit = null;
    var gapDot = 1e9;
    for (var j = 0; j < hits.length; j++) {
        if (hits[j] === restoreHit) continue;
        var dot2 = Math.abs(restoreHit.dir[0] * hits[j].dir[0] + restoreHit.dir[1] * hits[j].dir[1]);
        if (dot2 < gapDot) {
            gapDot = dot2;
            gapHit = hits[j];
        }
    }
    return { restore: restoreHit, gap: gapHit, dot: gapDot };
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

function MDUX_gapTools_findGapMarkerNear(gapLayer, center, tol) {
    if (!gapLayer || !center) return null;
    var best = null;
    var bestDist = (typeof tol === "number") ? tol : 9999;
    for (var i = 0; i < gapLayer.pathItems.length; i++) {
        try {
            var item = gapLayer.pathItems[i];
            if (!item.note || item.note.indexOf("MDUX_GAP:") !== 0) continue;
            var meta = MDUX_gapTools_parseGapMetadata(item.note);
            var pt = null;
            if (meta && typeof meta.x === "number" && typeof meta.y === "number") {
                pt = [meta.x, meta.y];
            } else if (item.pathPoints && item.pathPoints.length > 0) {
                pt = [item.pathPoints[0].anchor[0], item.pathPoints[0].anchor[1]];
            }
            if (!pt) continue;
            var dist = MDUX_gapTools_pointDistance(center, pt);
            if (dist <= bestDist) {
                bestDist = dist;
                best = item;
            }
        } catch (e) { }
    }
    return best;
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

function MDUX_gapTools_restoreDeletedSegment(seg, targetLayer, targetCompound) {
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
        if (targetCompound) {
            try { newPath.move(targetCompound, ElementPlacement.PLACEATEND); } catch (e) { }
        }
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
    try {
        if (!path || !cutBefore || !cutAfter) return null;
        if (!MDUX_gapTools_isValidPath(path)) {
            MDUX_gapTools_trace("SPLIT abort invalid path");
            return null;
        }
        var first = null, second = null;
        var targetLayer = parentLayer || null;
        var targetCompound = null;
        if (!targetLayer) {
            try { targetLayer = MDUX_gapTools_getParentLayer(path); } catch (e) { targetLayer = null; }
        }
        try { if (!targetLayer) targetLayer = path.layer; } catch (e) { }
        if (!MDUX_gapTools_isValidLayer(targetLayer)) {
            try { targetLayer = app.activeDocument.activeLayer; } catch (eAlt) { targetLayer = null; }
        }

        try { targetCompound = MDUX_gapTools_getParentCompound(path); } catch (e) { targetCompound = null; }
        if (targetCompound && targetCompound.isValid === false) targetCompound = null;

        var targetContainer = targetCompound || targetLayer;
        MDUX_gapTools_trace("SPLIT start segIdx=" + segIdx +
            " targetLayer=" + MDUX_gapTools_safeLayerName(targetLayer) +
            " compound=" + (targetCompound ? "yes" : "no"));

        var pts = null;
        try { pts = path.pathPoints; } catch (ePts) { pts = null; }
        if (!pts || pts.length < 2) {
            MDUX_gapTools_trace("SPLIT abort missing points");
            return null;
        }
        if (segIdx < 0 || segIdx >= pts.length - 1) {
            MDUX_gapTools_trace("SPLIT abort invalid segIdx=" + segIdx + " pts=" + pts.length);
            return null;
        }
        MDUX_gapTools_trace("SPLIT points ok count=" + pts.length);

        var firstPts = [];
        var secondPts = [];
        try {
            for (var i = 0; i <= segIdx; i++) {
                firstPts.push([pts[i].anchor[0], pts[i].anchor[1]]);
            }
            firstPts.push([cutBefore[0], cutBefore[1]]);
            secondPts.push([cutAfter[0], cutAfter[1]]);
            for (var j = segIdx + 1; j < pts.length; j++) {
                secondPts.push([pts[j].anchor[0], pts[j].anchor[1]]);
            }
        } catch (eBuild) {
            MDUX_gapTools_traceError("SPLIT build points error", eBuild);
            return null;
        }
        MDUX_gapTools_trace("SPLIT built points first=" + firstPts.length + " second=" + secondPts.length);

        function createPathWithStyle(points) {
            var newPath = null;
            try {
                if (!MDUX_gapTools_isValidLayer(targetLayer)) {
                    MDUX_gapTools_trace("SPLIT invalid targetLayer on create");
                    return null;
                }
                newPath = targetLayer.pathItems.add();
            } catch (eAdd) {
                MDUX_gapTools_traceError("SPLIT create path error", eAdd);
                return null;
            }
            try { newPath.setEntirePath(points); } catch (eSet) {
                MDUX_gapTools_traceError("SPLIT set path error", eSet);
            }
            try { newPath.filled = path.filled; } catch (eFill) { }
            try { newPath.stroked = path.stroked; } catch (eStroke) { }
            try { newPath.strokeWidth = path.strokeWidth; } catch (eStrokeW) { }
            try { newPath.strokeColor = path.strokeColor; } catch (eStrokeC) { }
            try { if (path.filled) newPath.fillColor = path.fillColor; } catch (eFillC) { }
            return newPath;
        }

        first = createPathWithStyle(firstPts);
        second = createPathWithStyle(secondPts);
        MDUX_gapTools_trace("SPLIT created paths first=" + (first ? "yes" : "no") + " second=" + (second ? "yes" : "no"));
        if (!first || !second) return null;

        if (targetCompound) {
            try { if (first.parent !== targetCompound) first.move(targetCompound, ElementPlacement.PLACEATEND); } catch (e) { MDUX_gapTools_traceError("SPLIT move first error", e); }
            try { if (second.parent !== targetCompound) second.move(targetCompound, ElementPlacement.PLACEATEND); } catch (e) { MDUX_gapTools_traceError("SPLIT move second error", e); }
        }

        try { path.remove(); } catch (e) { MDUX_gapTools_traceError("SPLIT remove original error", e); }
        return { first: first, second: second };
    } catch (eAll) {
        MDUX_gapTools_traceError("SPLIT fatal error", eAll);
        return null;
    }
}

function MDUX_healGapsInSelection() {
    var trace = [];
    var traceId = new Date().getTime();
    var traceTag = "HEAL-" + traceId;
    MDUX_gapTools_setTraceBuffer(trace);
    MDUX_gapTools_setTraceTag(traceTag);
    try {
        if (app.documents.length === 0) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
            return "No document open";
        }
        var doc = app.activeDocument;
        var selectedPaths = MDUX_gapTools_collectSelectedPaths(doc);
        selectedPaths = MDUX_gapTools_filterValidPaths(selectedPaths);
        if (!selectedPaths || selectedPaths.length < 1) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
            return "Select ductwork paths to heal gaps";
        }
        function traceMsg(msg) { MDUX_gapTools_trace(msg); }
        function fmtPt(pt) { return pt ? (pt[0].toFixed(1) + "," + pt[1].toFixed(1)) : "null"; }
        traceMsg("HEAL start id=" + traceId + " selected=" + selectedPaths.length + " build=" + MDUX_gapTools_buildTag);

        var bounds = MDUX_gapTools_getSelectionBounds(selectedPaths);
        var gapLayer = MDUX_gapTools_getGapLayer(doc, false);
        if (!gapLayer) {
            var traceFileEarly = MDUX_gapTools_finalizeTrace(traceTag, trace);
            return "No Gap Definitions layer (trace " + traceId + ")" + (traceFileEarly ? " file: " + traceFileEarly : "");
        }

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
        traceMsg("HEAL markers=" + markers.length);

        for (var m = 0; m < markers.length; m++) {
            var entry = markers[m];
            var marker = entry.marker;
            var info = entry.info;
            var center = info.center;
            var gapSize = info.gapSize;
            var dir = info.dir;
            var sourceLayer = info.sourceLayer;
            traceMsg("HEAL marker " + m + " center=" + fmtPt(center) + " sourceLayer=" + (sourceLayer || ""));

            var restored = null;
            var nearPath = MDUX_gapTools_findNearestSegment(selectedPaths, center, null, 25);
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
                        if (nearPath && nearPath.path && nearPath.path.layer) targetLayer = nearPath.path.layer;
                    }
                    if (targetLayer) {
                        var targetCompound = null;
                        try {
                            if (nearPath && nearPath.path) {
                                targetCompound = MDUX_gapTools_getParentCompound(nearPath.path);
                            }
                        } catch (e) { targetCompound = null; }
                        restored = MDUX_gapTools_restoreDeletedSegment(seg, targetLayer, targetCompound);
                        if (restored) {
                            restoredCount++;
                            restoredSegments.push(restored);
                            traceMsg("HEAL restored segment to layer=" + (targetLayer.name || "") + " compound=" + (targetCompound ? "yes" : "no"));
                        }
                    }
                }
            }

            var filled = false;
            var nearFilled = MDUX_gapTools_findNearestSegment(selectedPaths, center, sourceLayer, 4);
            if (nearFilled && nearFilled.dist <= 4) filled = true;

            if (!restored && !filled) {
                skippedMissing++;
                traceMsg("HEAL skipped (missing restore/filled) center=" + fmtPt(center));
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
            traceMsg("HEAL done center=" + fmtPt(center));
        }

        try { gapLayer.locked = gapLayerWasLocked; } catch (e) { }
        try { gapLayer.visible = gapLayerWasVisible; } catch (e) { }

        if (restoredSegments.length > 0) {
            var mergedTotal = 0;
            for (var rs = 0; rs < restoredSegments.length; rs++) {
                var restoredPath = restoredSegments[rs];
                try { if (restoredPath.isValid === false) continue; } catch (e) { }
                var candidates = [];
                var compound = null;
                try { compound = MDUX_gapTools_getParentCompound(restoredPath); } catch (e) { compound = null; }
                if (compound) {
                    candidates = MDUX_gapTools_collectOpenPathsFromCompound(compound);
                } else {
                    for (var sp = 0; sp < selectedPaths.length; sp++) {
                        var p = selectedPaths[sp];
                        if (!p) continue;
                        var exists = false;
                        for (var ci = 0; ci < candidates.length; ci++) {
                            if (candidates[ci] === p) { exists = true; break; }
                        }
                        if (!exists) candidates.push(p);
                    }
                    for (var rp = 0; rp < restoredSegments.length; rp++) {
                        var rpPath = restoredSegments[rp];
                        if (!rpPath) continue;
                        var exists2 = false;
                        for (var ci2 = 0; ci2 < candidates.length; ci2++) {
                            if (candidates[ci2] === rpPath) { exists2 = true; break; }
                        }
                        if (!exists2) candidates.push(rpPath);
                    }
                }
                if (candidates.length > 0) {
                    mergedTotal += MDUX_gapTools_mergePathWithNeighbors(restoredPath, candidates, 3.0, 0.985);
                }
            }
            traceMsg("HEAL merged=" + mergedTotal);
        }

        var traceFile = MDUX_gapTools_finalizeTrace(traceTag, trace);
        if (markers.length < 1) return "No gap markers near selection (trace " + traceId + ")" + (traceFile ? " file: " + traceFile : "");
        if (healedCount < 1 && skippedMissing > 0) {
            return "No gaps healed (missing deleted segments) (trace " + traceId + ")" + (traceFile ? " file: " + traceFile : "");
        }

        return "Healed " + healedCount + " gap(s), restored " + restoredCount + ", removed " + markerRemoved +
            " marker(s), patched " + patchedCount + ", removed " + ignoreRemoved + " ignore anchor(s)" +
            (skippedMissing > 0 ? ", skipped " + skippedMissing : "") + " (trace " + traceId + ")" +
            (traceFile ? " file: " + traceFile : "");
    } catch (e) {
        try { trace.push("ERROR: " + e); } catch (e2) { }
        var traceFile = "";
        try { traceFile = MDUX_gapTools_finalizeTrace(traceTag, trace); } catch (e3) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
        }
        return "Error: " + e + " (trace " + traceId + ")" + (traceFile ? " file: " + traceFile : "");
    }
}

function MDUX_recreateGapsInSelection() {
    var trace = [];
    var traceId = new Date().getTime();
    var traceTag = "RECREATE-" + traceId;
    MDUX_gapTools_setTraceBuffer(trace);
    MDUX_gapTools_setTraceTag(traceTag);
    try {
        if (app.documents.length === 0) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
            return "No document open";
        }
        var doc = app.activeDocument;
        var selectedPaths = MDUX_gapTools_collectSelectedPaths(doc);
        selectedPaths = MDUX_gapTools_filterValidPaths(selectedPaths);
        if (!selectedPaths || selectedPaths.length < 1) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
            return "Select ductwork paths to recreate gaps";
        }
        function traceMsg(msg) { MDUX_gapTools_trace(msg); }
        function fmtPt(pt) { return pt ? (pt[0].toFixed(1) + "," + pt[1].toFixed(1)) : "null"; }
        traceMsg("RECREATE start id=" + traceId + " selected=" + selectedPaths.length + " build=" + MDUX_gapTools_buildTag);

        var bounds = MDUX_gapTools_getSelectionBounds(selectedPaths);
        var gapLayer = MDUX_gapTools_getGapLayer(doc, true);
        if (!gapLayer) {
            var traceFileEarly = MDUX_gapTools_finalizeTrace(traceTag, trace);
            return "Unable to access Gap Definitions layer (trace " + traceId + ")" + (traceFileEarly ? " file: " + traceFileEarly : "");
        }

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

        var activePaths = MDUX_gapTools_filterValidPaths(selectedPaths.slice());
        var targets = [];
        var targetCenters = [];
        var restoredSegments = [];
        var restoredCount = 0;
        var ignoreRemoved = 0;
        var flippedCount = 0;
        traceMsg("RECREATE activePaths=" + activePaths.length);
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
            targetCenters.push(center);
        }
        traceMsg("RECREATE markers/anchors=" + targets.length);

        var intersectionCenters = MDUX_gapTools_collectIntersectionCenters(activePaths, 1);
        traceMsg("RECREATE intersections=" + intersectionCenters.length);
        if (intersectionCenters.length > 0) {
            for (var ic = 0; ic < intersectionCenters.length; ic++) {
                var interCenter = intersectionCenters[ic];
                var tooClose = false;
                for (var tc = 0; tc < targetCenters.length; tc++) {
                    if (MDUX_gapTools_pointDistance(interCenter, targetCenters[tc]) <= 8) {
                        tooClose = true;
                        break;
                    }
                }
                if (tooClose) continue;
                targets.push({ item: null, type: "intersection", center: interCenter, info: null });
                targetCenters.push(interCenter);
            }
        }
        traceMsg("RECREATE targets=" + targets.length);

        var recreated = 0;
        var skipped = 0;
        var skippedExisting = 0;

        for (var t = 0; t < targets.length; t++) {
            var target = targets[t];
            var center = target.center;
            var info = target.info || {};
            var sourceLayer = info.sourceLayer || null;
            traceMsg("RECREATE target " + t + " type=" + target.type + " center=" + fmtPt(center) + " sourceLayer=" + (sourceLayer || ""));

            var hits = MDUX_gapTools_findIntersectingSegmentsAtPoint(activePaths, center, 8);
            var pairInfo = MDUX_gapTools_chooseBestPairByDirection(hits);
            var hitA = pairInfo ? pairInfo.a : (hits.length > 0 ? hits[0] : null);
            var hitB = pairInfo ? pairInfo.b : (hits.length > 1 ? hits[1] : null);
            traceMsg("RECREATE hits=" + hits.length + (pairInfo ? (" pairDot=" + pairInfo.dot.toFixed(3)) : ""));

            var near = null;
            if (sourceLayer) {
                near = MDUX_gapTools_findNearestSegment(activePaths, center, sourceLayer, 25);
            }
            if (!near) near = hitA;
            if (!near) {
                near = MDUX_gapTools_findNearestSegment(activePaths, center, null, 25);
            }
            if (!near) { skipped++; continue; }

            var existingSeg = null;
            if (deletedLayer) {
                existingSeg = MDUX_gapTools_findDeletedSegmentNear(deletedLayer, center, 15);
                if (!existingSeg && near) {
                    existingSeg = MDUX_gapTools_findDeletedSegmentNear(deletedLayer, near.proj, 15);
                }
            }
            traceMsg("RECREATE existingSeg=" + (existingSeg ? "yes" : "no"));

            var gapHit = near;
            var wasFlip = false;

            if (existingSeg && hits.length > 0) {
                var segPts = null;
                try { segPts = existingSeg.pathPoints; } catch (e) { segPts = null; }
                if (!segPts || segPts.length < 2) { skippedExisting++; continue; }

                var segStart = segPts[0].anchor;
                var segEnd = segPts[segPts.length - 1].anchor;
                var ddx = segEnd[0] - segStart[0];
                var ddy = segEnd[1] - segStart[1];
                var dlen = Math.sqrt(ddx * ddx + ddy * ddy);
                var dirDel = (dlen > 0.001) ? [ddx / dlen, ddy / dlen] : [1, 0];
                var flipHits = MDUX_gapTools_chooseFlipHits(hits, dirDel);
                if (!flipHits || !flipHits.restore || !flipHits.gap) { skippedExisting++; continue; }
                if (flipHits.dot >= 0.98) { skippedExisting++; continue; }

                var restoreHit = flipHits.restore;
                gapHit = flipHits.gap;
                traceMsg("RECREATE flip restoreDot=" + flipHits.dot.toFixed(3));

                var restoreLayer = null;
                var restoreCompound = null;
                try {
                    restoreLayer = restoreHit.path.layer;
                    restoreCompound = MDUX_gapTools_getParentCompound(restoreHit.path);
                } catch (e) {
                    restoreLayer = null;
                    restoreCompound = null;
                }
                if (!restoreLayer && sourceLayer) {
                    try { restoreLayer = doc.layers.getByName(sourceLayer); } catch (e) { restoreLayer = null; }
                }

                var restored = null;
                if (restoreLayer) {
                    restored = MDUX_gapTools_restoreDeletedSegment(existingSeg, restoreLayer, restoreCompound);
                }
                if (!restored) { skippedExisting++; continue; }

                restoredCount++;
                restoredSegments.push(restored);
                ignoreRemoved += MDUX_gapTools_removeIgnoreAnchorsNear(doc, [[segStart[0], segStart[1]], [segEnd[0], segEnd[1]]], 12);
                flippedCount++;
                wasFlip = true;
                traceMsg("RECREATE flip restored to layer=" + (restoreLayer ? restoreLayer.name : "") + " compound=" + (restoreCompound ? "yes" : "no"));
            } else if (existingSeg) {
                skippedExisting++;
                continue;
            }

            if (target.type === "intersection" && hitA && hitB && !wasFlip) {
                var distA = Math.min(hitA.t, 1 - hitA.t);
                var distB = Math.min(hitB.t, 1 - hitB.t);
                if (Math.abs(distA - distB) > 0.001) {
                    gapHit = (distA >= distB) ? hitA : hitB;
                } else {
                    var lenA = MDUX_gapTools_getPathLength(hitA.path);
                    var lenB = MDUX_gapTools_getPathLength(hitB.path);
                    gapHit = (lenA <= lenB) ? hitA : hitB;
                }
            }

            if (!gapHit) { skipped++; continue; }
            var gapHitPath = null;
            try { gapHitPath = gapHit.path; } catch (eGap) {
                MDUX_gapTools_traceError("RECREATE gapHit path error", eGap);
                skipped++;
                continue;
            }
            if (!MDUX_gapTools_isValidPath(gapHitPath)) {
                gapHit = MDUX_gapTools_findNearestSegment(activePaths, center, sourceLayer, 25);
                gapHitPath = null;
                try { if (gapHit) gapHitPath = gapHit.path; } catch (eGap2) { gapHitPath = null; }
                if (!gapHit || !MDUX_gapTools_isValidPath(gapHitPath)) {
                    traceMsg("RECREATE gapHit invalid");
                    skipped++;
                    continue;
                }
                traceMsg("RECREATE gapHit refreshed");
            }

            var segLen = gapHit.segLen;
            if (!segLen || segLen < 0.01) { skipped++; continue; }

            var effectiveSourceLayer = sourceLayer;
            if (!effectiveSourceLayer || wasFlip || target.type === "intersection") {
                try { if (gapHitPath && gapHitPath.layer) effectiveSourceLayer = gapHitPath.layer.name; } catch (e) { }
            }

            var gapSize = info.gapSize;
            var isAutoSized = info.isAutoSized;
            if (!gapSize || gapSize <= 0) {
                var strokeWidth = null;
                try { strokeWidth = gapHitPath.strokeWidth; } catch (e) { strokeWidth = null; }
                gapSize = MDUX_gapTools_calculateAutoGapSize(strokeWidth);
                isAutoSized = true;
            }

            var dir = info.dir;
            if (!dir || dir.length < 2 || wasFlip || target.type === "intersection") {
                dir = gapHit.dir;
            }
            if (!dir || dir.length < 2) {
                var dx = gapHit.segEnd[0] - gapHit.segStart[0];
                var dy = gapHit.segEnd[1] - gapHit.segStart[1];
                var len = Math.sqrt(dx * dx + dy * dy);
                if (len < 0.01) { skipped++; continue; }
                dir = [dx / len, dy / len];
            }

            var ratio = gapSize / segLen;
            if (ratio >= 0.49) { skipped++; continue; }
            if (gapHit.t <= ratio || gapHit.t >= (1 - ratio)) { skipped++; continue; }

            var centerOnSeg = gapHit.proj;
            var cutBefore = [centerOnSeg[0] - gapSize * dir[0], centerOnSeg[1] - gapSize * dir[1]];
            var cutAfter = [centerOnSeg[0] + gapSize * dir[0], centerOnSeg[1] + gapSize * dir[1]];
            traceMsg("RECREATE gapHit len=" + segLen.toFixed(1) + " t=" + gapHit.t.toFixed(3) + " gapSize=" + gapSize.toFixed(2));
            var pathValid = false;
            try { pathValid = MDUX_gapTools_isValidPath(gapHitPath); } catch (eValid) {
                MDUX_gapTools_traceError("RECREATE gapHit path valid error", eValid);
                pathValid = false;
            }
            traceMsg("RECREATE gapHit path valid=" + pathValid);
            if (!pathValid) {
                traceMsg("RECREATE gapHit invalid after check");
                skipped++;
                continue;
            }

            if (!wasFlip && deletedLayer) {
                var existingSeg2 = MDUX_gapTools_findDeletedSegmentNear(deletedLayer, centerOnSeg, 15);
                if (existingSeg2) { skippedExisting++; continue; }
            }

            try {
                traceMsg("RECREATE step saveDeletedSegment");
                MDUX_gapTools_saveDeletedSegment(doc, cutBefore, cutAfter, effectiveSourceLayer || "");
            } catch (eSave) {
                MDUX_gapTools_traceError("RECREATE saveDeletedSegment error", eSave);
                skipped++;
                continue;
            }

            if (target.type === "gap") {
                if (target.item) {
                    traceMsg("RECREATE target item valid=" + MDUX_gapTools_isValidPath(target.item));
                }
                try {
                    traceMsg("RECREATE step createGapMarker");
                    MDUX_gapTools_createGapMarker(doc, gapLayer, centerOnSeg, gapSize, effectiveSourceLayer, dir, isAutoSized, target.item);
                } catch (eMarker) {
                    MDUX_gapTools_traceError("RECREATE createGapMarker error", eMarker);
                    skipped++;
                    continue;
                }
            } else {
                if (target.type === "patch") {
                    try { target.item.locked = false; } catch (e) { }
                    try { target.item.remove(); } catch (e) { }
                }
                var reuseMarker = null;
                if (wasFlip) {
                    reuseMarker = MDUX_gapTools_findGapMarkerNear(gapLayer, center, 12);
                }
                try {
                    traceMsg("RECREATE step createGapMarker");
                    MDUX_gapTools_createGapMarker(doc, gapLayer, centerOnSeg, gapSize, effectiveSourceLayer, dir, isAutoSized, reuseMarker);
                } catch (eMarker2) {
                    MDUX_gapTools_traceError("RECREATE createGapMarker error", eMarker2);
                    skipped++;
                    continue;
                }
            }

            if (ignoreLayer) {
                try {
                    traceMsg("RECREATE step createIgnoreAnchor");
                    MDUX_gapTools_createIgnoreAnchor(ignoreLayer, cutBefore);
                    MDUX_gapTools_createIgnoreAnchor(ignoreLayer, cutAfter);
                } catch (eIgnore) {
                    MDUX_gapTools_traceError("RECREATE createIgnoreAnchor error", eIgnore);
                    skipped++;
                    continue;
                }
            }

            var parentLayer = MDUX_gapTools_getParentLayer(gapHitPath);
            var split = null;
            try {
                traceMsg("RECREATE step splitPath");
                split = MDUX_gapTools_splitPathAtGap(gapHitPath, gapHit.segIdx, cutBefore, cutAfter, parentLayer);
            } catch (eSplit) {
                MDUX_gapTools_traceError("RECREATE splitPath error", eSplit);
                skipped++;
                continue;
            }
            if (split && split.first && split.second) {
                traceMsg("RECREATE split returned first=" + (split.first ? "yes" : "no") +
                    " second=" + (split.second ? "yes" : "no"));
                var removed = false;
                var nextActive = [];
                var gapHitPathValid = false;
                try { gapHitPathValid = MDUX_gapTools_isValidPath(gapHitPath); } catch (eGapValid) { gapHitPathValid = false; }
                for (var ap = 0; ap < activePaths.length; ap++) {
                    var candidate = null;
                    try { candidate = activePaths[ap]; } catch (eAp0) { candidate = null; }
                    if (!candidate) continue;
                    if (!MDUX_gapTools_isValidPath(candidate)) continue;
                    var isMatch = false;
                    if (gapHitPathValid) {
                        try { isMatch = (candidate === gapHitPath); } catch (eAp) {
                            MDUX_gapTools_traceError("RECREATE activePaths remove error", eAp);
                            isMatch = false;
                        }
                    }
                    if (isMatch) {
                        removed = true;
                        continue;
                    }
                    nextActive.push(candidate);
                }
                activePaths = nextActive;
                traceMsg("RECREATE activePaths removed=" + removed);
                try {
                    if (MDUX_gapTools_isValidPath(split.first) && split.first.pathPoints && split.first.pathPoints.length >= 2) {
                        activePaths.push(split.first);
                        traceMsg("RECREATE split first pushed");
                    } else {
                        traceMsg("RECREATE split first invalid");
                    }
                } catch (eFirst) {
                    MDUX_gapTools_traceError("RECREATE split first error", eFirst);
                }
                try {
                    if (MDUX_gapTools_isValidPath(split.second) && split.second.pathPoints && split.second.pathPoints.length >= 2) {
                        activePaths.push(split.second);
                        traceMsg("RECREATE split second pushed");
                    } else {
                        traceMsg("RECREATE split second invalid");
                    }
                } catch (eSecond) {
                    MDUX_gapTools_traceError("RECREATE split second error", eSecond);
                }
                recreated++;
                traceMsg("RECREATE split ok");
            } else {
                skipped++;
                traceMsg("RECREATE split failed");
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

        if (restoredSegments.length > 0) {
            var mergedTotal = 0;
            for (var rs = 0; rs < restoredSegments.length; rs++) {
                var restoredPath = restoredSegments[rs];
                try { if (restoredPath.isValid === false) continue; } catch (e) { }
                var candidates = [];
                var compound = null;
                try { compound = MDUX_gapTools_getParentCompound(restoredPath); } catch (e) { compound = null; }
                if (compound) {
                    candidates = MDUX_gapTools_collectOpenPathsFromCompound(compound);
                } else {
                    for (var sp = 0; sp < selectedPaths.length; sp++) {
                        var p = selectedPaths[sp];
                        if (!p) continue;
                        var exists = false;
                        for (var ci = 0; ci < candidates.length; ci++) {
                            if (candidates[ci] === p) { exists = true; break; }
                        }
                        if (!exists) candidates.push(p);
                    }
                    for (var rp = 0; rp < restoredSegments.length; rp++) {
                        var rpPath = restoredSegments[rp];
                        if (!rpPath) continue;
                        var exists2 = false;
                        for (var ci2 = 0; ci2 < candidates.length; ci2++) {
                            if (candidates[ci2] === rpPath) { exists2 = true; break; }
                        }
                        if (!exists2) candidates.push(rpPath);
                    }
                }
                if (candidates.length > 0) {
                    mergedTotal += MDUX_gapTools_mergePathWithNeighbors(restoredPath, candidates, 3.0, 0.985);
                }
            }
            traceMsg("RECREATE merged=" + mergedTotal);
        }

        var traceFile = MDUX_gapTools_finalizeTrace(traceTag, trace);
        if (targets.length < 1) return "No gap markers or intersections near selection (trace " + traceId + ")" +
            (traceFile ? " file: " + traceFile : "");

        var msg = "Recreated " + recreated + " gap(s)";
        if (flippedCount > 0) msg += ", flipped " + flippedCount;
        if (restoredCount > 0) msg += ", restored " + restoredCount;
        if (ignoreRemoved > 0) msg += ", removed " + ignoreRemoved + " ignore anchor(s)";
        if (skippedExisting > 0) msg += ", skipped " + skippedExisting + " existing";
        if (skipped > 0) msg += ", skipped " + skipped;
        return msg + " (trace " + traceId + ")" + (traceFile ? " file: " + traceFile : "");
    } catch (e) {
        try { trace.push("ERROR: " + e); } catch (e2) { }
        var traceFile = "";
        try { traceFile = MDUX_gapTools_finalizeTrace(traceTag, trace); } catch (e3) {
            MDUX_gapTools_clearTraceTag();
            MDUX_gapTools_clearTraceBuffer();
        }
        return "Error: " + e + " (trace " + traceId + ")" + (traceFile ? " file: " + traceFile : "");
    }
}
