/**
 * Python Geometry Bridge for Magic Ductwork
 *
 * Calls the Python geometry engine for heavy computational operations.
 * This provides 50-500x speedup over pure ExtendScript implementations.
 *
 * Usage:
 *   var result = PythonBridge.findConnections(pathsArray);
 *   var groups = PythonBridge.buildGroups(pathsArray);
 */

var PythonBridge = (function() {

    // Configuration
    var PYTHON_EXE = "python";  // Assumes python is in PATH
    var ENGINE_SCRIPT = "geometry_engine.py";
    var DEBUG = true;

    /**
     * Get the path to the Python engine script
     */
    function getEnginePath() {
        try {
            // Try to get from global JSX folder
            var jsxFolder = $.global.MDUX_JSX_FOLDER;
            if (jsxFolder) {
                var pythonFolder = new Folder(jsxFolder.parent.fsName + "/python");
                if (pythonFolder.exists) {
                    return pythonFolder.fsName + "/" + ENGINE_SCRIPT;
                }
            }

            // Fallback: relative to current script
            var currentScript = new File($.fileName);
            var pythonFolder = new Folder(currentScript.parent.parent.fsName + "/python");
            if (pythonFolder.exists) {
                return pythonFolder.fsName + "/" + ENGINE_SCRIPT;
            }

            // Last resort: hardcoded development path
            return "e:/Work/Work/Custom Sketchup, Illustrator and Photoshop Scripts and Extensions/Illustrator/Extensions/Magic-Ductwork-Panel/python/" + ENGINE_SCRIPT;
        } catch (e) {
            return null;
        }
    }

    /**
     * Get a temporary file path for JSON data exchange
     */
    function getTempFilePath(suffix) {
        var tempFolder = Folder.temp;
        var timestamp = new Date().getTime();
        return tempFolder.fsName + "/mdux_" + suffix + "_" + timestamp + ".json";
    }

    /**
     * Write JSON data to a file
     */
    function writeJsonFile(filePath, data) {
        var file = new File(filePath);
        file.encoding = "UTF-8";
        file.open('w');
        file.write(JSON.stringify(data));
        file.close();
        return file.exists;
    }

    /**
     * Read JSON data from a file
     */
    function readJsonFile(filePath) {
        var file = new File(filePath);
        if (!file.exists) {
            return null;
        }
        file.encoding = "UTF-8";
        file.open('r');
        var content = file.read();
        file.close();

        try {
            return JSON.parse(content);
        } catch (e) {
            if (DEBUG) addDebug("[PYBRIDGE] JSON parse error: " + e);
            return null;
        }
    }

    /**
     * Delete a temporary file
     */
    function deleteTempFile(filePath) {
        try {
            var file = new File(filePath);
            if (file.exists) {
                file.remove();
            }
        } catch (e) {
            // Ignore cleanup errors
        }
    }

    /**
     * Convert Illustrator PathItem to JSON-serializable format
     */
    function pathToJson(pathItem) {
        var points = [];
        try {
            var pp = pathItem.pathPoints;
            for (var i = 0; i < pp.length; i++) {
                var anchor = pp[i].anchor;
                points.push({
                    x: anchor[0],
                    y: anchor[1]
                });
            }
        } catch (e) {
            // Invalid path
        }
        return { points: points };
    }

    /**
     * Convert array of PathItems to JSON format
     */
    function pathsToJson(pathItems) {
        var paths = [];
        for (var i = 0; i < pathItems.length; i++) {
            var p = pathItems[i];
            if (!p) continue;

            var pathData = pathToJson(p);
            pathData.id = i;

            // Store reference info for later
            try {
                pathData.layerName = p.layer ? p.layer.name : null;
            } catch (e) {
                pathData.layerName = null;
            }

            paths.push(pathData);
        }
        return paths;
    }

    /**
     * Execute Python geometry engine
     * @param {string} operation - The operation to run (find_connections, build_groups, etc.)
     * @param {Array} pathItems - Array of Illustrator PathItems
     * @param {Object} params - Optional parameters for the operation
     * @returns {Object} Result from Python engine, or null on error
     */
    function executePython(operation, pathItems, params) {
        var startTime = new Date().getTime();

        if (DEBUG) addDebug("[PYBRIDGE] Starting " + operation + " with " + pathItems.length + " paths");

        // Get engine path
        var enginePath = getEnginePath();
        if (!enginePath) {
            addDebug("[PYBRIDGE] ERROR: Could not find geometry_engine.py");
            return null;
        }

        // Convert paths to JSON
        var pathsJson = pathsToJson(pathItems);

        // Prepare input data
        var inputData = {
            paths: pathsJson,
            params: params || {}
        };

        // Write input file
        var inputPath = getTempFilePath("input");
        var outputPath = getTempFilePath("output");

        if (!writeJsonFile(inputPath, inputData)) {
            addDebug("[PYBRIDGE] ERROR: Failed to write input file");
            return null;
        }

        // Build command
        // Windows: use cmd /c to handle piping
        var cmd = 'cmd /c "' + PYTHON_EXE + ' "' + enginePath + '" ' + operation +
                  ' < "' + inputPath + '" > "' + outputPath + '" 2>&1"';

        if (DEBUG) addDebug("[PYBRIDGE] Executing: " + cmd);

        // Execute Python
        try {
            app.system(cmd);
        } catch (e) {
            addDebug("[PYBRIDGE] ERROR executing Python: " + e);
            deleteTempFile(inputPath);
            return null;
        }

        // Read output
        var result = readJsonFile(outputPath);

        // Cleanup
        deleteTempFile(inputPath);
        deleteTempFile(outputPath);

        var elapsed = new Date().getTime() - startTime;

        if (result) {
            if (DEBUG) addDebug("[PYBRIDGE] " + operation + " completed in " + elapsed + "ms (Python: " + (result.time_ms || "?") + "ms)");

            if (result.error) {
                addDebug("[PYBRIDGE] Python error: " + result.error);
                return null;
            }
        } else {
            addDebug("[PYBRIDGE] ERROR: No result from Python");
        }

        return result;
    }

    /**
     * Find all connections between paths
     * @param {Array} pathItems - Array of PathItems to analyze
     * @param {number} maxDist - Maximum distance for connection detection (default: 10)
     * @returns {Object} { connections: [...], ignored_anchors: [...] }
     */
    function findConnections(pathItems, maxDist, tTolerance) {
        return executePython('find_connections', pathItems, { max_dist: maxDist || 10, t_tolerance: tTolerance });
    }

    /**
     * Build groups of connected paths
     * @param {Array} pathItems - Array of PathItems to group
     * @param {number} maxDist - Maximum distance for connection detection
     * @returns {Object} { groups: [[indices], ...], connections: [...] }
     */
    function buildGroups(pathItems, maxDist) {
        return executePython('build_groups', pathItems, { max_dist: maxDist || 10 });
    }

    /**
     * Detect all intersections between paths
     * @param {Array} pathItems - Array of PathItems
     * @returns {Object} { intersections: [...] }
     */
    function detectIntersections(pathItems) {
        return executePython('detect_intersections', pathItems, {});
    }

    /**
     * Orthogonalize paths (snap + make horizontal/vertical)
     * This is the main performance bottleneck - Python is 100x+ faster
     * @param {Array} pathItems - Array of PathItems to orthogonalize
     * @param {number} snapThreshold - Snap distance threshold (default: 5)
     * @returns {Object} { paths: [{id, points}], iterations, time_ms }
     */
    function orthogonalize(pathItems, snapThreshold) {
        return executePython('orthogonalize', pathItems, {
            snap_threshold: snapThreshold || 5,
            steep_min: 17,
            steep_max: 70
        });
    }

    /**
     * Check if Python bridge is available
     */
    function isAvailable() {
        var enginePath = getEnginePath();
        if (!enginePath) return false;

        var engineFile = new File(enginePath);
        return engineFile.exists;
    }

    // Public API
    return {
        findConnections: findConnections,
        buildGroups: buildGroups,
        detectIntersections: detectIntersections,
        orthogonalize: orthogonalize,
        isAvailable: isAvailable,

        // Expose utilities for debugging
        _getEnginePath: getEnginePath,
        _pathsToJson: pathsToJson
    };

})();

// Make available globally
$.global.PythonBridge = PythonBridge;
