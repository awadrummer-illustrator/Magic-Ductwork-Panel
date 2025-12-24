
/**
 * Export Utilities for Magic Ductwork
 */

// ========================================
// CONFIGURATION
// ========================================

var EXPORT_LAYERS_DUCTWORK = [
    // Ductwork Pieces
    "Thermostats",
    "Units",
    "Secondary Exhaust Registers",
    "Exhaust Registers",
    "Orange Register",
    "Rectangular Registers",
    "Square Registers",
    "Circular Registers",
    // Ductwork Lines
    "Green Ductwork",
    "Light Green Ductwork",
    "Blue Ductwork",
    "Orange Ductwork",
    "Light Orange Ductwork",
    "Thermostat Lines",
    // Frame
    "Frame"
];

var EXPORT_LAYERS_FLOORPLAN = [
    "Render",
    "Frame"
];

// ========================================
// HELPER FUNCTIONS
// ========================================

var MDUX_logBuffer = [];

function MDUX_log(msg) {
    MDUX_logBuffer.push(msg);
}

function MDUX_getLog() {
    return MDUX_logBuffer.join("\n");
}

function MDUX_getFloorplanItem(doc) {
    // Helper to find best item in a specific layer
    function getBestFromLayer(layerName, requirePNG) {
        try {
            var layer = doc.layers.getByName(layerName);
            var best = null;
            var maxArea = 0;
            for (var i = 0; i < layer.placedItems.length; i++) {
                var item = layer.placedItems[i];
                try {
                    if (!item.file) continue; // Must have linked file
                    
                    // Check for PNG if required
                    if (requirePNG) {
                        var fName = item.file.name.toLowerCase();
                        if (fName.slice(-4) !== ".png") continue;
                    }

                    var area = item.width * item.height;
                    // Pick the largest one. If equal, keeps the first one found (since > not >=).
                    if (area > maxArea) {
                        maxArea = area;
                        best = item;
                    }
                } catch(e) {}
            }
            return best;
        } catch(e) {
            return null;
        }
    }

    // Priority 1: Render Layer (Highest priority for naming) - MUST be PNG
    var item = getBestFromLayer("Render", true);
    if (item) return item;

    // Priority 2: Floorplan Layer
    item = getBestFromLayer("Floorplan", false);
    if (item) return item;

    // Priority 3: Ductwork Reference Layer
    item = getBestFromLayer("Ductwork Reference", false);
    if (item) return item;

    // Fallback: Largest in document
    var bestItem = null;
    var maxArea = 0;
    for (var i = 0; i < doc.placedItems.length; i++) {
        var item = doc.placedItems[i];
        try {
            if (!item.file) continue;
            var area = item.width * item.height;
            if (area > maxArea) {
                maxArea = area;
                bestItem = item;
            }
        } catch(e) {}
    }
    return bestItem;
}

function MDUX_setLayerVisibility(doc, layerNamesToKeepVisible) {
    var originalStates = {};
    
    // Store original states and hide all
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        originalStates[layer.name] = layer.visible;
        layer.visible = false;
    }

    // Show requested layers
    for (var j = 0; j < layerNamesToKeepVisible.length; j++) {
        try {
            var layer = doc.layers.getByName(layerNamesToKeepVisible[j]);
            layer.visible = true;
        } catch (e) {
            // Layer might not exist, ignore
        }
    }

    return originalStates;
}

function MDUX_restoreLayerVisibility(doc, originalStates) {
    for (var name in originalStates) {
        if (originalStates.hasOwnProperty(name)) {
            try {
                var layer = doc.layers.getByName(name);
                layer.visible = originalStates[name];
            } catch (e) {}
        }
    }
}

function MDUX_ensureFolderExists(folderPath) {
    var folder = new Folder(folderPath);
    if (!folder.exists) {
        folder.create();
    }
    return folder;
}

function MDUX_exportPNG(doc, file, scalePercent, artboardRect) {
    var options = new ExportOptionsPNG24();
    options.antiAliasing = true;
    options.transparency = true;
    options.artBoardClipping = true;
    options.horizontalScale = scalePercent;
    options.verticalScale = scalePercent;

    // If we have a specific rect, we might need to adjust the artboard
    // But for now, let's assume the user wants the active artboard or we set it
    
    doc.exportFile(file, ExportType.PNG24, options);
}

// ========================================
// MAIN EXPORT FUNCTION
// ========================================

function MDUX_performExport(exportType, overwrite, versionSuffix) {
    MDUX_logBuffer = []; // Clear buffer at start
    MDUX_log("MDUX_performExport called with: " + exportType + ", overwrite=" + overwrite + ", version=" + versionSuffix);
    
    try {
        if (app.documents.length === 0) {
            MDUX_log("No document open.");
            return JSON.stringify({ ok: false, message: "No document open.", log: MDUX_getLog() });
        }
        
        var doc = app.activeDocument;
        var floorplanItem = MDUX_getFloorplanItem(doc);
        
        if (!floorplanItem) {
            MDUX_log("Could not find Floorplan reference image (must be a linked PlacedItem).");
            return JSON.stringify({ ok: false, message: "Could not find Floorplan reference image (must be a linked PlacedItem).", log: MDUX_getLog() });
        }

        // Get reference file info
        var refFile = null;
        try {
            refFile = floorplanItem.file;
        } catch (e) {
            MDUX_log("Floorplan item is not linked to a file: " + e.message);
            return JSON.stringify({ ok: false, message: "Floorplan item is not linked to a file.", log: MDUX_getLog() });
        }

        if (!refFile) {
            MDUX_log("Floorplan item has no source file.");
            return JSON.stringify({ ok: false, message: "Floorplan item has no source file.", log: MDUX_getLog() });
        }

        MDUX_log("Found floorplan file: " + refFile.fsName);

        var refFileName = decodeURI(refFile.name);
        var refNameBase = refFileName.substring(0, refFileName.lastIndexOf("."));
        var parentFolder = refFile.parent;

        // Calculate Dimensions and Scale
        var itemWidth = floorplanItem.width; // Points
        var itemHeight = floorplanItem.height; // Points
        
        MDUX_log("Floorplan dimensions: " + itemWidth + "x" + itemHeight);

        // Target dimensions
        // If Width > Height: Width = 4000px
        // If Height > Width: Height = 2160px
        
        var targetWidth4K, targetHeight4K;
        var scaleFactor4K;

        if (itemWidth > itemHeight) {
            targetWidth4K = 4000;
            scaleFactor4K = (targetWidth4K / itemWidth) * 100; // Percentage
        } else {
            targetHeight4K = 2160;
            scaleFactor4K = (targetHeight4K / itemHeight) * 100; // Percentage
        }

        var scaleFactor2K = scaleFactor4K / 2;
        
        MDUX_log("Scale Factor 4K: " + scaleFactor4K + "%");

        // Define Suffixes and Layers
        var typeSuffix = "";
        var layersToExport = [];
        
        if (exportType === "DUCTWORK") {
            typeSuffix = " Ductwork";
            layersToExport = EXPORT_LAYERS_DUCTWORK.slice(); // Copy array
        } else if (exportType === "FLOORPLAN") {
            typeSuffix = " Floorplan"; 
            layersToExport = EXPORT_LAYERS_FLOORPLAN.slice(); // Copy array
        }

        // Versioning
        var versionStr = "";
        if (versionSuffix) {
            versionStr = " V" + versionSuffix;
        }

        // Construct Filenames
        var name4K = refNameBase + typeSuffix + versionStr + "_4K.png";
        var name2K = refNameBase + typeSuffix + versionStr + "_2K.png";

        // Folders
        // User requested a single folder named "2K, 4K and 8K"
        var exportFolderName = "2K, 4K and 8K";
        var exportFolder = MDUX_ensureFolderExists(parentFolder.fsName + "/" + exportFolderName);

        var file4K = new File(exportFolder.fsName + "/" + name4K);
        var file2K = new File(exportFolder.fsName + "/" + name2K);
        
        MDUX_log("Target 4K file: " + file4K.fsName);

        // Check existence
        if (!overwrite && !versionSuffix) {
            if (file4K.exists || file2K.exists) {
                MDUX_log("Files exist, asking for confirmation.");
                return JSON.stringify({ 
                    ok: true, 
                    status: "CONFIRM_OVERWRITE", 
                    message: "Files already exist. Overwrite?",
                    existingFiles: [file4K.fsName, file2K.fsName],
                    log: MDUX_getLog()
                });
            }
        }

        // Prepare Artboard
        // We want to export exactly the area of the floorplan item.
        // Best way: Resize active artboard to match floorplan item, then restore.
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var originalArtboardRect = doc.artboards[activeArtboardIndex].artboardRect;
        
        // Set artboard to floorplan bounds
        // Rect is [left, top, right, bottom]
        // Item bounds: visibleBounds or geometricBounds? geometricBounds is safer for image.
        var bounds = floorplanItem.geometricBounds; 
        doc.artboards[activeArtboardIndex].artboardRect = bounds;

        // Set Visibility
        var originalVis = MDUX_setLayerVisibility(doc, layersToExport);

        try {
            // Export 4K
            MDUX_log("Exporting 4K...");
            MDUX_exportPNG(doc, file4K, scaleFactor4K);
            
            // Export 2K
            MDUX_log("Exporting 2K...");
            MDUX_exportPNG(doc, file2K, scaleFactor2K);
            
        } catch (e) {
            MDUX_log("Export failed: " + e.message);
            return JSON.stringify({ ok: false, message: "Export failed: " + e.message, log: MDUX_getLog() });
        } finally {
            // Restore
            MDUX_restoreLayerVisibility(doc, originalVis);
            doc.artboards[activeArtboardIndex].artboardRect = originalArtboardRect;
        }

        MDUX_log("Export successful.");
        return JSON.stringify({ ok: true, status: "SUCCESS", message: "Exported successfully to " + parentFolder.fsName, log: MDUX_getLog() });

    } catch (globalErr) {
        MDUX_log("CRITICAL ERROR: " + globalErr.message);
        return JSON.stringify({ ok: false, message: "Critical Error: " + globalErr.message, log: MDUX_getLog() });
    }
}
