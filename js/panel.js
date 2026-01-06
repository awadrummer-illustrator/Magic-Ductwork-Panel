(function () {
    'use strict';

    const csInterface = new CSInterface();
    const processBtn = document.getElementById('process-btn');
    const processEmoryBtn = document.getElementById('process-emory-btn');
    const processStatus = document.getElementById('process-status');
    const revertBtn = document.getElementById('revert-ortho-btn');
    const revertStatus = document.getElementById('revert-status');
    const clearRotationMetadataBtn = document.getElementById('clear-rotation-metadata-btn');
    const clearRotationMetadataStatus = document.getElementById('clear-rotation-metadata-status');
    // const applyIgnoreBtn = document.getElementById('apply-ignore-btn'); // Removed - button no longer exists
    // const ignoreStatus = document.getElementById('ignore-status'); // Removed - status no longer exists
    const reloadBtn = document.getElementById('reload-btn');
    const debugStatus = document.getElementById('debug-status');
    const skipOrthoOption = document.getElementById('skip-ortho-option');
    const rotationInput = document.getElementById('rotation-input');
    const getAngleBtn = document.getElementById('get-angle-btn');
    const clearRotationBtn = document.getElementById('clear-rotation-btn');
    const skipAllBranchesOption = document.getElementById('skip-all-branches-option');
    const skipFinalOption = document.getElementById('skip-final-option');
    const skipRegisterRotationOption = document.getElementById('skip-register-rotation-option');
    const createRegisterWiresOption = document.getElementById('create-register-wires-option');
    const rotate90Btn = document.getElementById('rotate-90-btn');
    const rotate45Btn = document.getElementById('rotate-45-btn');
    const rotate180Btn = document.getElementById('rotate-180-btn');
    const rotateCustomBtn = document.getElementById('rotate-custom-btn');
    const customRotationInput = document.getElementById('custom-rotation-input');
    const scaleSlider = document.getElementById('scale-slider');
    const scaleLabel = document.getElementById('scale-label');
    const scaleInput = document.getElementById('scale-input');
    const applyScaleBtn = document.getElementById('apply-scale-btn');
    const resetScaleBtn = document.getElementById('reset-scale-btn');
    const selectionStatus = document.getElementById('selection-status');
    const isolatePartsBtn = document.getElementById('isolate-parts-btn');
    const isolateLinesBtn = document.getElementById('isolate-lines-btn');
    const unlockDuctworkBtn = document.getElementById('unlock-ductwork-btn');
    const createLayersBtn = document.getElementById('create-layers-btn');
    const layerStatus = document.getElementById('layer-status');
    const importStylesBtn = document.getElementById('import-styles-btn');
    const importStatus = document.getElementById('import-status');
    const exportDuctworkBtn = document.getElementById('export-ductwork-btn');
    const reexportFloorplanBtn = document.getElementById('reexport-floorplan-btn');
    const exportStatus = document.getElementById('export-status');

    // Document Scale Controls (read-only anchor display)
    const docScaleInput = document.getElementById('doc-scale-input');
    const getDocScaleBtn = document.getElementById('get-doc-scale-btn');
    // Legacy: Document scaling controls disabled to prevent desync issues
    // const setDocScale100Btn = document.getElementById('set-doc-scale-100-btn');
    // const applyDocScaleBtn = document.getElementById('apply-doc-scale-btn');

    // Transform Each Controls
    let teScaleInput = document.getElementById('te-scale');
    let teRotateInput = document.getElementById('te-rotate');
    let teScaleSlider = document.getElementById('te-scale-slider');
    let teRotateSlider = document.getElementById('te-rotate-slider');
    let transformEachBtn = document.getElementById('transform-each-btn');
    let teResetOriginalBtn = document.getElementById('te-reset-original-btn');
    let teLiveOption = document.getElementById('te-live-option');

    // Reset/Normalize Controls
    const resetStrokesBtn = document.getElementById('reset-strokes-btn');
    const resetPartsScaleBtn = document.getElementById('reset-parts-scale-btn');

    // Collapsible Section Controls
    const docScaleToggle = document.getElementById('doc-scale-toggle');
    const docScaleSection = document.getElementById('doc-scale-section');

    let scaleDebounce = null;
    let bridgeReloaded = false;
    let skipOrthoRefreshTimer = null;
    const bridgePath = (function () {
        var root = csInterface.getSystemPath(CSInterface.SystemPath.EXTENSION);
        return root.replace(/\\/g, '/') + '/jsx/panel-bridge.jsx';
    })();;

    function evalScript(script) {
        return new Promise((resolve, reject) => {
            csInterface.evalScript(script, (result) => {
                if (typeof result === 'undefined' || result === null) {
                    resolve('');
                } else {
                    resolve(result);
                }
            });
        });
    }

    function setProcessStatus(message, isError) {
        if (!processStatus) return;
        processStatus.textContent = message || '';
        processStatus.classList.toggle('error', !!isError);
    }

    function setIgnoreStatus(message, isError) {
        // ignoreStatus element removed - function disabled
    }

    function setRevertStatus(message, isError) {
        if (!revertStatus) return;
        revertStatus.textContent = message || '';
        revertStatus.classList.toggle('error', !!isError);
    }

    function setClearRotationMetadataStatus(message, isError) {
        if (!clearRotationMetadataStatus) return;
        clearRotationMetadataStatus.textContent = message || '';
        clearRotationMetadataStatus.classList.toggle('error', !!isError);
    }

    function setSelectionStatus(message, isError) {
        const el = document.getElementById('selection-status');
        const debugEl = document.getElementById('debug-status');

        if (el) {
            el.textContent = message || '';
            el.classList.toggle('error', !!isError);
            el.style.display = 'block'; // Force display
        }

        if (debugEl) {
            // Also show in footer for redundancy
            debugEl.textContent = (isError ? '[ERR] ' : '') + (message || '');
            debugEl.title = message || ''; // Tooltip for full text
        }
    }

    function setLayerStatus(message, isError) {
        if (!layerStatus) return;
        layerStatus.textContent = message || '';
        layerStatus.classList.toggle('error', !!isError);
    }

    function setImportStatus(message, isError) {
        if (!importStatus) return;
        importStatus.textContent = message || '';
        importStatus.classList.toggle('error', !!isError);
    }

    function setExportStatus(message, isError) {
        if (!exportStatus) return;
        exportStatus.textContent = message || '';
        exportStatus.classList.toggle('error', !!isError);
    }

    function normaliseResult(value) {
        if (!value) return { ok: true, value: '' };
        if (typeof value === 'string' && value.indexOf('ERROR:') === 0) {
            return { ok: false, value: value.substring(6) };
        }
        return { ok: true, value: value };
    }

    function escapeForExtendScript(str) {
        return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    }

    async function ensureBridgeLoaded() {
        // Check if bridge is already loaded to avoid unnecessary reloads
        if (bridgeReloaded) {
            return;
        }
        const escapedPath = escapeForExtendScript(bridgePath);
        const loadScript = '(function(){' +
            'delete $.global.MDUX_JSX_FOLDER;' +  // Force clear stale cached folder path
            'delete $.global.MDUX;' +  // Force clear stale MDUX namespace to ensure fresh initialization
            '$.global.MDUX_LAST_BRIDGE_PATH = "' + escapedPath + '";' +
            'try { $.evalFile("' + escapedPath + '"); return "OK"; } ' +
            'catch (e) { $.global.MDUX_LAST_BRIDGE_ERROR = e.toString(); return "ERROR:" + e; }' +
            '})()';
        const loadResult = await evalScript(loadScript);
        if (typeof loadResult === 'string' && loadResult.indexOf('ERROR:') === 0) {
            const msg = loadResult.substring(6);
            debugStatus.textContent = 'Bridge load failed: ' + msg;
            throw new Error(msg);
        }
        bridgeReloaded = true;
        debugStatus.textContent = 'Bridge ready: ' + bridgePath.replace(/\\/g, '/');
    }

    function scheduleSkipOrthoRefresh() {
        console.log('[JS] scheduleSkipOrthoRefresh called');
        evalScript('MDUX_debugLog("[JS] scheduleSkipOrthoRefresh called")');
        if (skipOrthoRefreshTimer) clearTimeout(skipOrthoRefreshTimer);
        skipOrthoRefreshTimer = setTimeout(() => {
            console.log('[JS] scheduleSkipOrthoRefresh timeout fired, calling refresh functions');
            evalScript('MDUX_debugLog("[JS] scheduleSkipOrthoRefresh timeout fired")');
            refreshSkipOrthoState().catch(() => { });
            refreshRotationOverrideState().catch(() => { });
            refreshSelectionTransformState().catch(() => { });
        }, 150);
    }

    async function refreshRotationOverrideState() {
        if (!rotationInput) return;

        // NEVER update if user is typing in the input
        var isFocused = document.activeElement === rotationInput;
        if (isFocused) return;

        try {
            await ensureBridgeLoaded();
        } catch (e) {
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_rotationStateBridge()'));
        if (!result.ok) {
            if (result.value) {
                debugStatus.textContent = 'Rotation state error: ' + result.value;
            }
            return;
        }
        var summary = null;
        try { summary = result.value ? JSON.parse(result.value) : null; } catch (eParse) { summary = null; }
        if (!summary) return;

        var nextPlaceholder = 'Leave blank to skip';

        if (summary.reason === 'no-document') {
            rotationInput.value = '';
            rotationInput.dataset.autoValue = '';
            rotationInput.dataset.multi = 'false';
            rotationInput.placeholder = 'No document';
            return;
        }

        if (summary.reason === 'no-selection') {
            rotationInput.value = '';
            rotationInput.dataset.autoValue = '';
            rotationInput.dataset.multi = 'false';
            rotationInput.placeholder = nextPlaceholder;
            return;
        }

        if (!summary.available) {
            return;
        }

        var formatted = summary.formatted || '';
        rotationInput.value = formatted;
        rotationInput.dataset.autoValue = formatted;

        if (summary.count > 1) {
            rotationInput.dataset.multi = 'true';
            nextPlaceholder = 'Mixed rotations';
        } else {
            rotationInput.dataset.multi = 'false';
        }

        rotationInput.placeholder = nextPlaceholder;
    }

    async function refreshSkipOrthoState() {
        if (!skipOrthoOption) return;
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            skipOrthoOption.indeterminate = false;
            skipOrthoOption.checked = false;
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_skipOrthoStateBridge()'));
        if (!result.ok) {
            skipOrthoOption.indeterminate = false;
            skipOrthoOption.checked = false;
            debugStatus.textContent = 'Skip-ortho state error: ' + result.value;
            return;
        }
        var state = null;
        try { state = result.value ? JSON.parse(result.value) : null; } catch (eParse) { state = null; }
        if (!state || state.available === false) {
            skipOrthoOption.indeterminate = false;
            skipOrthoOption.checked = false;
            skipOrthoOption.disabled = !!(state && state.reason === 'no-document');
            return;
        }
        skipOrthoOption.disabled = false;
        if (state.mixed) {
            skipOrthoOption.indeterminate = true;
            skipOrthoOption.checked = false;
        } else if (state.hasNote) {
            skipOrthoOption.indeterminate = false;
            skipOrthoOption.checked = true;
        } else {
            skipOrthoOption.indeterminate = false;
            skipOrthoOption.checked = false;
        }
    }

    async function handleProcessClick() {
        processBtn.disabled = true;
        revertBtn.disabled = true;
        setProcessStatus('Preparing ductwork run…');
        setRevertStatus('');

        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setProcessStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            processBtn.disabled = false;
            revertBtn.disabled = false;
            return;
        }
        let rotationValue = null;
        const rotationText = rotationInput.value.trim();
        const autoValue = (rotationInput.dataset.autoValue || '').trim();
        // Skip rotation override if user hasn't changed the auto-populated value
        const isAutoPopulated = rotationText === autoValue && rotationText.length > 0;
        if (rotationText && !isAutoPopulated) {
            rotationValue = parseFloat(rotationText);
            if (!isFinite(rotationValue)) {
                setProcessStatus('Rotation override must be a valid number.', true);
                processBtn.disabled = false;
                revertBtn.disabled = false;
                return;
            }
        }

        const skipAutoCarveOption = document.getElementById('skip-auto-carve-option');
        const options = {
            action: 'process',
            skipAllBranchSegments: !!skipAllBranchesOption.checked,
            skipFinalRegisterSegment: !!skipFinalOption.checked,
            skipRegisterRotation: !(skipRegisterRotationOption && skipRegisterRotationOption.checked),
            skipAutoCarve: !!(skipAutoCarveOption && skipAutoCarveOption.checked)
        };

        if (!skipOrthoOption.indeterminate) {
            options.skipOrtho = !!skipOrthoOption.checked;
        }
        if (rotationValue !== null) {
            options.rotationOverride = rotationValue;
        }

        const payload = JSON.stringify(options).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        const prep = normaliseResult(await evalScript("MDUX_prepareProcessBridge('" + payload + "')"));
        if (!prep.ok) {
            setProcessStatus('Error: ' + prep.value, true);
            processBtn.disabled = false;
            revertBtn.disabled = false;
            return;
        }

        setProcessStatus('Running ductwork script…');
        const result = normaliseResult(await evalScript('MDUX_runMagicDuctwork()'));
        if (result.ok) {
            setProcessStatus('Magic Ductwork completed.');
            debugStatus.textContent = 'Process completed';
        } else {
            setProcessStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Process failed: ' + result.value;
        }
        processBtn.disabled = false;
        revertBtn.disabled = false;
        scheduleSkipOrthoRefresh();

        // Auto-copy debug log to clipboard and write to file after processing
        try {
            console.log('[PANEL] Calling MDUX_getDebugLog()...');
            const debugLog = await evalScript('MDUX_getDebugLog()');
            console.log('[PANEL] MDUX_getDebugLog returned ' + (debugLog ? debugLog.length : 0) + ' chars');
            if (debugLog && debugLog.length > 0) {
                // Use textarea + execCommand for CEP compatibility
                const ta = document.createElement('textarea');
                ta.value = debugLog;
                ta.style.position = 'fixed';
                ta.style.left = '-9999px';
                document.body.appendChild(ta);
                ta.select();
                const copyResult = document.execCommand('copy');
                document.body.removeChild(ta);
                console.log('[PANEL] Debug log copied to clipboard: ' + copyResult + ' (' + debugLog.length + ' chars)');
            }
        } catch (clipErr) {
            console.log('[PANEL] Failed to copy debug log to clipboard:', clipErr);
        }
    }

    async function handleProcessEmoryClick() {
        processEmoryBtn.disabled = true;
        setProcessStatus('Running Emory ductwork processing…');

        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setProcessStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            processEmoryBtn.disabled = false;
            return;
        }

        const createWires = !!createRegisterWiresOption.checked;
        const result = normaliseResult(await evalScript('MDUX_runEmoryDuctwork(' + createWires + ')'));
        if (result.ok) {
            setProcessStatus('Ready.');
            debugStatus.textContent = 'Emory process completed';
        } else {
            setProcessStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Emory process failed: ' + result.value;
        }
        processEmoryBtn.disabled = false;
    }

    async function rotateSelection(angle) {
        if (!isFinite(angle)) {
            setSelectionStatus('Rotation value must be numeric.', true);
            return;
        }
        setSelectionStatus('Rotating selection ' + angle + '°…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setSelectionStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_rotateSelectionBridge(' + angle + ')'));
        if (!result.ok) {
            setSelectionStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Rotate failed: ' + result.value;
            return;
        }
        let stats = null;
        try { stats = result.value ? JSON.parse(result.value) : null; } catch (e) { stats = null; }
        if (stats && typeof stats.total === 'number') {
            if (stats.total === 0) {
                setSelectionStatus('Select units/registers to rotate.', true);
                debugStatus.textContent = 'Rotate: no eligible items';
            } else if (stats.rotated > 0) {
                setSelectionStatus('Rotated ' + stats.rotated + ' item(s).' + (stats.skipped ? ' Skipped ' + stats.skipped + '.' : ''));
                debugStatus.textContent = 'Rotate result: rotated ' + stats.rotated + ', skipped ' + (stats.skipped || 0);
            } else {
                setSelectionStatus('No eligible items were rotated.', true);
                debugStatus.textContent = 'Rotate: no items rotated';
            }
        } else {
            setSelectionStatus(result.value || 'Rotation complete.');
            debugStatus.textContent = 'Rotate result: ' + (result.value || 'OK');
        }
        scheduleSkipOrthoRefresh();
    }

    async function applyScale(percent) {
        setSelectionStatus('Scaling selection to ' + percent + '%…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setSelectionStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }

        // First, reset graphic styles on ductwork lines to ensure proper scaling
        try {
            await evalScript('MDUX.resetDuctworkLineStyles()');
        } catch (e) {
            // Non-fatal - continue with scaling even if reset fails
        }

        const result = normaliseResult(await evalScript('MDUX_scaleSelectionBridge(' + percent + ')'));
        if (!result.ok) {
            setSelectionStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Scale failed: ' + result.value;
            return;
        }
        let stats = null;
        try { stats = result.value ? JSON.parse(result.value) : null; } catch (e) { stats = null; }
        if (stats && typeof stats.total === 'number') {
            if (stats.total === 0) {
                setSelectionStatus('Select ductwork, registers, units, or thermostats to scale.', true);
                debugStatus.textContent = 'Scale: no eligible items';
            } else if (stats.scaled > 0) {
                setSelectionStatus('Scaled ' + stats.scaled + ' item(s).' + (stats.skipped ? ' Skipped ' + stats.skipped + '.' : ''));
                debugStatus.textContent = 'Scale result: scaled ' + stats.scaled + ', skipped ' + (stats.skipped || 0);
            } else {
                setSelectionStatus('No eligible items were scaled.', true);
                debugStatus.textContent = 'Scale: no items scaled';
            }
        } else {
            setSelectionStatus(result.value || 'Scaling complete.');
            debugStatus.textContent = 'Scale result: ' + (result.value || 'OK');
        }
        scheduleSkipOrthoRefresh();
    }

    async function resetScale() {
        setSelectionStatus('Resetting selection scale…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setSelectionStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_resetScaleBridge()'));
        if (!result.ok) {
            setSelectionStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Reset scale failed: ' + result.value;
            return;
        }
        let stats = null;
        try { stats = result.value ? JSON.parse(result.value) : null; } catch (e) { stats = null; }
        if (stats && typeof stats.total === 'number') {
            if (stats.total === 0) {
                setSelectionStatus('No eligible items selected to reset.', true);
                debugStatus.textContent = 'Reset scale: no eligible items';
            } else if (stats.reset > 0) {
                setSelectionStatus('Restored scale on ' + stats.reset + ' item(s).' + (stats.skipped ? ' Skipped ' + stats.skipped + '.' : ''));
                debugStatus.textContent = 'Reset scale result: reset ' + stats.reset + ', skipped ' + (stats.skipped || 0);
            } else {
                setSelectionStatus('Unable to restore scale for the selection.', true);
                debugStatus.textContent = 'Reset scale: nothing reset';
            }
        } else {
            setSelectionStatus(result.value || 'Scale reset complete.');
            debugStatus.textContent = 'Reset scale result: ' + (result.value || 'OK');
        }
        scheduleSkipOrthoRefresh();
    }

    function onScaleInput() {
        if (!scaleSlider || !scaleLabel) return;
        var value = parseFloat(scaleSlider.value);
        if (!isFinite(value)) return;
        scaleLabel.textContent = value + '%';
        if (scaleDebounce) {
            clearTimeout(scaleDebounce);
        }
        scaleDebounce = setTimeout(function () {
            applyScale(value);
        }, 400);
    }

    async function isolate(action) {
        let cmd;
        let startMessage = 'Applying layer isolation…';
        if (action === 'parts') cmd = 'MDUX_isolatePartsBridge()';
        else if (action === 'lines') cmd = 'MDUX_isolateDuctworkBridge()';
        else if (action === 'unlock') cmd = 'MDUX_unlockDuctworkBridge()';
        else if (action === 'create') {
            cmd = 'MDUX_createLayersBridge()';
            startMessage = 'Ensuring standard ductwork layers…';
        } else return;

        setLayerStatus(startMessage);
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setLayerStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }
        const result = normaliseResult(await evalScript(cmd));
        if (!result.ok) {
            setLayerStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Layer command failed: ' + result.value;
        } else {
            setLayerStatus(result.value || 'Layer operation completed.');
            debugStatus.textContent = 'Layer command result: ' + (result.value || 'OK');
        }
        scheduleSkipOrthoRefresh();
    }

    async function importGraphicStyles() {
        setImportStatus('Importing graphic styles…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setImportStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            debugStatus.textContent = 'Bridge load failed: ' + (e && e.message ? e.message : e);
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_importGraphicStylesBridge()'));
        if (!result.ok) {
            setImportStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Import styles failed: ' + result.value;
            return;
        }
        setImportStatus(result.value || 'Graphic styles imported.');
        debugStatus.textContent = 'Import styles result: ' + (result.value || 'OK');
    }

    async function applyIgnore() {
        // applyIgnoreBtn.disabled = true; // Button removed
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setIgnoreStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            // applyIgnoreBtn.disabled = false; // Button removed
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_applyIgnoreBridge()'));
        if (!result.ok) {
            setIgnoreStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Ignore apply failed: ' + result.value;
        } else {
            let stats = null;
            try { stats = result.value ? JSON.parse(result.value) : null; } catch (e) { stats = null; }
            if (stats && typeof stats.total === 'number') {
                if (stats.error) {
                    setIgnoreStatus(stats.error, true);
                    debugStatus.textContent = 'Ignore apply error: ' + stats.error;
                } else if (stats.total === 0) {
                    if (stats.reason === 'no-selection') {
                        setIgnoreStatus('Select duct lines or ductwork parts to add to ignored layer.', true);
                        debugStatus.textContent = 'Ignore apply: no selection';
                    } else {
                        setIgnoreStatus('No eligible items found in the selection.', true);
                        debugStatus.textContent = 'Ignore apply: no eligible items';
                    }
                } else if (stats.added > 0 || stats.moved > 0) {
                    const parts = [];
                    if (stats.added > 0) parts.push(stats.added + ' anchor(s)');
                    if (stats.moved > 0) parts.push(stats.moved + ' part(s) moved');
                    const skippedFragment = stats.skipped ? ' Skipped ' + stats.skipped + '.' : '';
                    setIgnoreStatus('Ignored ' + parts.join(', ') + '.' + skippedFragment);
                    debugStatus.textContent = 'Ignore apply result: added ' + (stats.added || 0) + ', moved ' + (stats.moved || 0) + ', skipped ' + (stats.skipped || 0);
                } else {
                    setIgnoreStatus('Unable to process selection.', true);
                    debugStatus.textContent = 'Ignore apply: unable to process';
                }
            } else {
                setIgnoreStatus(result.value || 'Ignore markers created.');
                debugStatus.textContent = 'Ignore apply result: ' + (result.value || 'OK');
            }
        }
        // applyIgnoreBtn.disabled = false; // Button removed
        scheduleSkipOrthoRefresh();
    }

    async function handleRevertPreOrtho() {
        revertBtn.disabled = true;
        setRevertStatus('Restoring original geometry…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setRevertStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            revertBtn.disabled = false;
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_revertPreOrthoBridge()'));
        if (!result.ok) {
            setRevertStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Revert failed: ' + result.value;
        } else {
            var stats = null;
            try {
                stats = result.value ? JSON.parse(result.value) : null;
            } catch (e) {
                stats = null;
            }
            if (stats && typeof stats.total === 'number') {
                if (stats.total === 0) {
                    if (stats.reason === 'no-selection') {
                        setRevertStatus('Select one or more duct lines to revert.', true);
                        debugStatus.textContent = 'Revert: no selection';
                    } else {
                        setRevertStatus('No stored pre-orthogonalization data found on the selection.', true);
                        debugStatus.textContent = 'Revert: no stored data';
                    }
                } else if (stats.reverted > 0) {
                    setRevertStatus('Reverted ' + stats.reverted + ' of ' + stats.total + ' path(s) to their pre-orthogonalized state.');
                    debugStatus.textContent = 'Revert result: reverted ' + stats.reverted + ' of ' + stats.total;
                } else {
                    setRevertStatus('Selection contains no stored pre-orthogonalization data.', true);
                    debugStatus.textContent = 'Revert: nothing reverted';
                }
            } else {
                setRevertStatus(result.value || 'Revert completed.');
                debugStatus.textContent = 'Revert result: ' + (result.value || 'OK');
            }
        }
        revertBtn.disabled = false;
        scheduleSkipOrthoRefresh();
    }

    async function handleClearRotationMetadata() {
        clearRotationMetadataBtn.disabled = true;
        setClearRotationMetadataStatus('Clearing rotation metadata…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setClearRotationMetadataStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            clearRotationMetadataBtn.disabled = false;
            return;
        }
        const result = normaliseResult(await evalScript('MDUX_clearRotationMetadataBridge()'));
        if (!result.ok) {
            setClearRotationMetadataStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Clear rotation metadata failed: ' + result.value;
        } else {
            var count = 0;
            try {
                count = parseInt(result.value, 10);
            } catch (e) {
                count = 0;
            }
            if (count === 0) {
                setClearRotationMetadataStatus('No paths selected or no rotation metadata found.', true);
                debugStatus.textContent = 'Clear rotation metadata: nothing cleared';
            } else {
                setClearRotationMetadataStatus('Cleared rotation metadata from ' + count + ' path(s).');
                debugStatus.textContent = 'Clear rotation metadata: cleared ' + count + ' paths';
            }
        }
        clearRotationMetadataBtn.disabled = false;
        scheduleSkipOrthoRefresh();
    }

    async function handleGetAngle() {
        setProcessStatus('Getting angle from selected line…');
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            setProcessStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }
        const resultStr = await evalScript('MDUX_getSelectedLineAngleBridge()');
        const result = normaliseResult(resultStr);

        if (!result.ok) {
            setProcessStatus(result.value, true);
            debugStatus.textContent = 'Get angle failed: ' + result.value;
            return;
        }

        try {
            const data = JSON.parse(result.value);
            if (data.ok && typeof data.angle === 'number') {
                rotationInput.value = data.angle.toString();
                rotationInput.dataset.autoValue = '';
                rotationInput.dataset.multi = 'false';
                setProcessStatus(data.message || ('Angle set to ' + data.angle + '°'));
                debugStatus.textContent = 'Angle retrieved: ' + data.angle + '°';
            } else {
                setProcessStatus(data.message || 'Failed to get angle', true);
                debugStatus.textContent = 'Get angle failed: ' + (data.message || 'Unknown error');
            }
        } catch (e) {
            setProcessStatus('Error parsing angle result', true);
            debugStatus.textContent = 'Get angle parse error: ' + e;
        }
    }

    async function refreshDocScale() {
        try {
            await ensureBridgeLoaded();
            const scale = await evalScript('MDUX_getDocumentScale()');
            if (scale && isFinite(parseFloat(scale))) {
                docScaleInput.value = scale;
            }
        } catch (e) {
            console.error("Failed to refresh doc scale:", e);
        }
    }

    async function handleSetDocScale100() {
        try {
            await ensureBridgeLoaded();
            const result = await evalScript('MDUX_setDocumentScale(100)');
            if (result === 'OK') {
                docScaleInput.value = '100';
                setSelectionStatus('Baseline set to 100%.');
            } else {
                setSelectionStatus('Error: ' + result, true);
            }
        } catch (e) {
            setSelectionStatus('Error: ' + e.message, true);
        }
    }

    async function handleApplyDocScale() {
        const targetPercent = parseFloat(docScaleInput.value);
        if (isNaN(targetPercent) || targetPercent <= 0) {
            setSelectionStatus('Enter a valid positive scale percentage.', true);
            return;
        }
        setSelectionStatus('Scaling entire document to ' + targetPercent + '%...');
        try {
            await ensureBridgeLoaded();
            const resultStr = await evalScript('MDUX_applyScaleToFullDocument(' + targetPercent + ')');
            const result = JSON.parse(resultStr);
            if (result.ok) {
                setSelectionStatus(result.message);
                if (result.stats) {
                    debugStatus.textContent = 'Doc scale: ' + result.stats.parts + ' parts, ' + result.stats.lines + ' lines';
                }
            } else {
                setSelectionStatus('Error: ' + result.message, true);
            }
        } catch (e) {
            setSelectionStatus('Error: ' + e.message, true);
        }
    }

    async function handleExport(type) {
        setExportStatus('Exporting...', false);
        console.log("Starting export for type: " + type);

        try {
            // Initial attempt (overwrite=false, version=null)
            let resultStr = await evalScript(`MDUX_performExport("${type}", false, null)`);
            console.log("Export result string: ", resultStr);

            if (!resultStr) {
                throw new Error("No response from Illustrator. Check debug.log.");
            }

            let result;
            try {
                result = JSON.parse(resultStr);
            } catch (e) {
                throw new Error("Invalid JSON response: " + resultStr);
            }

            if (!result.ok) {
                if (result.log) {
                    console.log("Export Error Log:\n" + result.log);
                }
                setExportStatus(result.message, true);
                return;
            }

            if (result.status === "CONFIRM_OVERWRITE") {
                // Ask user
                const shouldOverwrite = confirm(result.message || "Files already exist. Overwrite them?");

                if (shouldOverwrite) {
                    // Retry with overwrite=true
                    resultStr = await evalScript(`MDUX_performExport("${type}", true, null)`);
                    console.log("Overwrite result string: ", resultStr);
                    result = JSON.parse(resultStr);
                } else {
                    // Ask for version
                    const version = prompt("Enter version suffix (e.g. '1' for V1):");
                    if (version) {
                        // Retry with version
                        resultStr = await evalScript(`MDUX_performExport("${type}", false, "${version}")`);
                        console.log("Version result string: ", resultStr);
                        result = JSON.parse(resultStr);
                    } else {
                        setExportStatus("Export cancelled.", false);
                        return;
                    }
                }
            }

            if (result.ok) {
                if (result.log) {
                    console.log("Export Success Log:\n" + result.log);
                }
                setExportStatus(result.message, false);
            } else {
                if (result.log) {
                    console.log("Export Error Log:\n" + result.log);
                }
                setExportStatus(result.message, true);
            }

        } catch (e) {
            console.error("Export error:", e);
            setExportStatus("Error: " + e.message, true);
        }
    }

    // Transform Each State
    let teIsBusy = false;
    let teDragActive = false;
    let teTransformAppliedInDrag = false;
    let teNextPayload = null; // {scale, rotate, undoPrevious}

    // Track start values for the current drag session
    let teDragStartScale = 100;
    let teDragStartRotate = 0;

    // Track committed values (where the slider was left after last drag)
    // We need this because the slider value is absolute (e.g. 110), but we need to calculate
    // the factor relative to the START of the drag.
    // Actually, we can just read the slider value on mousedown.

    async function processTransformQueue() {
        if (teIsBusy) return;
        if (!teNextPayload) return;

        teIsBusy = true;

        // Capture current payload
        const payload = teNextPayload;
        teNextPayload = null; // Clear it, so we can catch new updates

        // Update status for feedback
        // Debugging: Show start/current values to diagnose scaling issues
        // setSelectionStatus(`Live: Scale ${payload.scale.toFixed(1)}%, Rot ${payload.rotate.toFixed(1)}°`, false);
        // setSelectionStatus(`Live: ${payload.scale.toFixed(1)}% (Start:${teDragStartScale}->Cur:${teScaleSlider.value})`, false);

        try {
            // Add a timeout race to prevent hanging if Illustrator doesn't respond
            const transformPromise = evalScript(`MDUX_transformEach(${payload.scale}, ${payload.rotate}, ${payload.undoPrevious})`);
            const timeoutPromise = new Promise(resolve => setTimeout(() => resolve("TIMEOUT"), 1000));

            const result = await Promise.race([transformPromise, timeoutPromise]);

            if (result === "TIMEOUT") {
                setSelectionStatus("Transform timed out", true);
            } else {
                // If successful, mark that we have applied a transform in this drag session
                if (teDragActive) {
                    teTransformAppliedInDrag = true;
                }

                // DEBUG: Show the result message from JSX
                try {
                    const resObj = JSON.parse(result);
                    if (resObj && resObj.message) {
                        setSelectionStatus(resObj.message, false);
                    } else {
                        setSelectionStatus("No msg: " + result, false);
                    }
                } catch (e) {
                    setSelectionStatus("Parse err: " + result, true);
                }
            }
        } catch (e) {
            console.error("Transform error:", e);
            setSelectionStatus("Error: " + e.message, true);
        } finally {
            teIsBusy = false;
            // Check if more accumulated while we were busy
            if (teNextPayload) {
                // IMPORTANT: If we are processing a queued item, it MUST undo the previous one
                // if a transform was applied.
                if (teTransformAppliedInDrag) {
                    teNextPayload.undoPrevious = true;
                }
                processTransformQueue();
            }
        }
    }

    function handleLiveTransform() {
        // Check if Live is enabled
        if (teLiveOption && !teLiveOption.checked) return;

        // Ensure elements are found
        if (!teScaleSlider || !teRotateSlider) return;

        const currentScale = parseFloat(teScaleSlider.value);
        const currentRotate = parseFloat(teRotateSlider.value);

        if (isNaN(currentScale) || isNaN(currentRotate)) return;

        // Calculate factor/delta relative to DRAG START
        // If dragStartScale is 0 (safety), use 100
        const startS = teDragStartScale || 100;

        // Factor: If start was 100, current 110 -> factor 110/100*100 = 110.
        // If start was 110, current 120 -> factor 120/110*100 = 109.09.
        // This factor is what we send to JSX to apply to the object's state AT DRAG START.
        // Wait, if we undo, we revert to state AT DRAG START (actually state before last transform).
        // So yes, we want to apply the transform that takes us from START to CURRENT.

        const factor = (currentScale / startS) * 100;
        const deltaRot = currentRotate - teDragStartRotate;

        // Determine if we should undo the previous step
        // We undo if we have already applied a transform in this specific drag session
        const undoPrevious = teTransformAppliedInDrag;

        teNextPayload = {
            scale: factor,
            rotate: deltaRot,
            undoPrevious: undoPrevious
        };

        processTransformQueue();
    }

    function resetTransformControls(resetValues = true) {
        if (resetValues) {
            if (teScaleSlider) teScaleSlider.value = 100;
            if (teScaleInput) teScaleInput.value = 100;
            if (teRotateSlider) teRotateSlider.value = 0;
            if (teRotateInput) teRotateInput.value = 0;
        }

        teDragActive = false;
        teTransformAppliedInDrag = false;
        teNextPayload = null;
        teDragStartScale = 100;
        teDragStartRotate = 0;
    }

    async function handleTransformEach() {
        // If Live is ON, the button just resets the controls (commits the change)
        if (teLiveOption && teLiveOption.checked) {
            resetTransformControls(true);
            setSelectionStatus("Transformation committed.", false);
        } else {
            // If Live is OFF, apply the current values
            const s = parseFloat(teScaleInput.value) || 100;
            const r = parseFloat(teRotateInput.value) || 0;

            if (s === 100 && r === 0) {
                setSelectionStatus("No changes to apply.", false);
                return;
            }

            setSelectionStatus("Transforming...", false);
            try {
                await evalScript(`MDUX_transformEach(${s}, ${r}, false)`);
                setSelectionStatus("Transformation applied.", false);
                resetTransformControls(true);
                // Refresh selection state to load the metadata we just saved
                await refreshSelectionTransformState();
            } catch (e) {
                setSelectionStatus("Error: " + e.message, true);
            }
        }
    }

    async function handleResetOriginal() {
        setSelectionStatus("Resetting to original...", false);
        try {
            await evalScript('MDUX_resetTransforms()');
            setSelectionStatus("Reset complete.", false);
            resetTransformControls(true);
        } catch (e) {
            setSelectionStatus("Error: " + e.message, true);
        }
    }

    async function refreshSelectionTransformState() {
        console.log('[JS] refreshSelectionTransformState called, teDragActive=', teDragActive, 'teIsBusy=', teIsBusy);
        if (teDragActive || teIsBusy) {
            await evalScript('MDUX_debugLog("[JS] refreshSelectionTransformState skipped: teDragActive=" + ' + teDragActive + ' + ", teIsBusy=" + ' + teIsBusy + ')');
            return;
        }

        await evalScript('MDUX_debugLog("[JS] refreshSelectionTransformState proceeding...")');

        try {
            await ensureBridgeLoaded();
        } catch (e) {
            console.error('Bridge load failed in refreshSelectionTransformState:', e);
            await evalScript('MDUX_debugLog("[JS] Bridge load failed in refreshSelectionTransformState: ' + (e ? e.message || e.toString() : 'unknown') + '")');
            return;
        }

        try {
            const raw = await evalScript('MDUX_getSelectionTransformState()');
            console.log('[JS] refreshSelectionTransformState raw:', raw);
            if (!raw) {
                await evalScript('MDUX_debugLog("[JS] MDUX_getSelectionTransformState returned null/empty")');
                return;
            }
            const res = JSON.parse(raw);
            console.log('[JS] refreshSelectionTransformState parsed:', res);

            const statusEl = document.getElementById('selection-status');
            if (statusEl) statusEl.textContent = '';

            if (res.ok && res.count > 0) {
                let statusMsg = [];

                // Only update scale if user is not actively typing in the scale input
                const scaleInputHasFocus = document.activeElement === teScaleInput;
                if (!scaleInputHasFocus) {
                    if (res.mixedScale) {
                        console.log('[JS] Mixed scale detected');
                        teScaleInput.value = '';
                        teScaleInput.placeholder = 'Mixed';
                        teScaleSlider.value = 100;
                        statusMsg.push('Multiple different scales in selection');
                    } else {
                        console.log('[JS] Setting scale to:', res.scale);
                        teScaleInput.value = res.scale;
                        teScaleInput.placeholder = '';
                        teScaleSlider.value = res.scale;
                        console.log('[JS] Scale slider value now:', teScaleSlider.value);
                    }
                } else {
                    console.log('[JS] Skipping scale update - input has focus');
                }

                // Only update rotation if user is not actively typing in the rotation input
                const rotateInputHasFocus = document.activeElement === teRotateInput;
                if (!rotateInputHasFocus) {
                    if (res.mixedRotation) {
                        console.log('[JS] Mixed rotation detected');
                        teRotateInput.value = '';
                        teRotateInput.placeholder = 'Mixed';
                        teRotateSlider.value = 0;
                        statusMsg.push('Multiple different rotations in selection');
                    } else {
                        console.log('[JS] Setting rotation to:', res.rotation);
                        teRotateInput.value = res.rotation;
                        teRotateInput.placeholder = '';
                        teRotateSlider.value = res.rotation;
                        console.log('[JS] Rotation slider value now:', teRotateSlider.value);
                    }
                } else {
                    console.log('[JS] Skipping rotation update - input has focus');
                }

                if (statusEl && statusMsg.length > 0) {
                    statusEl.textContent = statusMsg.join(' • ');
                }
            } else {
                console.log('[JS] No selection or not ok, resetting to defaults');
                // No selection or no tagged items
                teScaleInput.value = 100;
                teScaleInput.placeholder = '';
                teScaleSlider.value = 100;
                teRotateInput.value = 0;
                teRotateInput.placeholder = '';
                teRotateSlider.value = 0;
            }
        } catch (e) {
            console.error('Refresh transform state failed:', e);
        }
    }

    // Helper to handle drag start/end globally
    async function handleDragStart() {
        // Load current state from metadata BEFORE setting teDragActive
        // (refreshSelectionTransformState skips if teDragActive is true)
        await refreshSelectionTransformState();

        // Now set drag active
        teDragActive = true;
        teTransformAppliedInDrag = false;

        // Capture start values from both sliders (now updated from metadata)
        teDragStartScale = parseFloat(teScaleSlider.value) || 100;
        teDragStartRotate = parseFloat(teRotateSlider.value) || 0;

        // Listen for global mouseup to end the session
        window.addEventListener('mouseup', handleDragEnd, { once: true });
    }

    function handleDragEnd() {
        teDragActive = false;
        // The transformation is already applied during 'input' events.
        // We do NOT reset controls here so the user can see where they left it.
        teTransformAppliedInDrag = false;
    }

    function initCollapsibleSections() {
        const collapsibleHeaders = document.querySelectorAll('.collapsible-header');
        collapsibleHeaders.forEach(header => {
            header.addEventListener('click', function() {
                const section = this.closest('.collapsible');
                const content = section.querySelector('.collapsible-content');
                const icon = this.querySelector('.collapse-icon');

                if (content.style.display === 'none') {
                    content.style.display = 'block';
                    icon.textContent = '▲';
                    section.classList.add('expanded');
                } else {
                    content.style.display = 'none';
                    icon.textContent = '▼';
                    section.classList.remove('expanded');
                }
            });
        });
    }

    function attachListeners() {
        if (processBtn) processBtn.addEventListener('click', handleProcessClick);
        if (processEmoryBtn) processEmoryBtn.addEventListener('click', handleProcessEmoryClick);
        if (revertBtn) revertBtn.addEventListener('click', handleRevertPreOrtho);
        if (clearRotationMetadataBtn) clearRotationMetadataBtn.addEventListener('click', handleClearRotationMetadata);
        if (getAngleBtn) getAngleBtn.addEventListener('click', handleGetAngle);
        if (clearRotationBtn && rotationInput) {
            clearRotationBtn.addEventListener('click', () => {
                rotationInput.value = '';
                rotationInput.dataset.autoValue = '';
                rotationInput.dataset.multi = 'false';
                rotationInput.placeholder = 'Leave blank to skip';
                setProcessStatus('');
            });
        }
        if (rotationInput) {
            rotationInput.addEventListener('input', () => {
                rotationInput.dataset.autoValue = '';
                rotationInput.dataset.multi = 'false';
            });
            // Handle Enter key to apply rotation
            rotationInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    const val = parseFloat(rotationInput.value);
                    if (!isNaN(val)) {
                        console.log('[ROTATION] Enter pressed, applying rotation: ' + val);
                        rotateSelection(val);
                    }
                }
            });
        }

        // Document Scale listeners (read-only - only refresh button active)
        if (getDocScaleBtn) getDocScaleBtn.addEventListener('click', refreshDocScale);
        // Legacy: Document scaling controls disabled to prevent desync issues
        // if (setDocScale100Btn) setDocScale100Btn.addEventListener('click', handleSetDocScale100);
        // if (applyDocScaleBtn) applyDocScaleBtn.addEventListener('click', handleApplyDocScale);

        // New Mutual Exclusivity Logic
        const orthoToggles = [skipOrthoOption, skipAllBranchesOption, skipFinalOption];
        const orthoGrid = document.getElementById('ortho-toggle-grid');

        function updateOrthoState(changedInput) {
            if (changedInput && changedInput.checked) {
                orthoToggles.forEach(t => {
                    if (t !== changedInput) t.checked = false;
                });
            }

            // Update visual state
            if (orthoGrid) {
                const anyChecked = orthoToggles.some(t => t.checked);
                if (anyChecked) {
                    orthoGrid.classList.add('has-selection');
                } else {
                    orthoGrid.classList.remove('has-selection');
                }
            }
        }

        orthoToggles.forEach(t => {
            if (t) {
                t.addEventListener('change', () => updateOrthoState(t));
            }
        });

        // Initial state check
        updateOrthoState(null);
        if (rotate90Btn) rotate90Btn.addEventListener('click', () => rotateSelection(90));
        if (rotate45Btn) rotate45Btn.addEventListener('click', () => rotateSelection(45));
        if (rotate180Btn) rotate180Btn.addEventListener('click', () => rotateSelection(180));
    }

    if (teScaleSlider && teScaleInput) {
        teScaleSlider.addEventListener('mousedown', handleDragStart);
        // Backup: change event fires on commit (release)
        // teScaleSlider.addEventListener('change', () => resetTransformControls(true)); // REMOVED

        teScaleSlider.addEventListener('input', () => {
            teScaleInput.value = teScaleSlider.value;
            handleLiveTransform();
        });

        teScaleInput.addEventListener('change', () => {
            let val = parseFloat(teScaleInput.value);
            if (isNaN(val)) return;

            // For text input, we treat it as a mini-session
            // We need a start value. Use current slider value as start?
            // Or assume start was 100 relative to current state?
            // If user types 150, they mean 150% of current state? Or 150% absolute?
            // Usually absolute.
            // But our logic is relative.
            // Let's assume they mean "Apply 150% scale".
            // So start=100, current=150. Factor=150.

            teDragStartScale = 100;
            teDragStartRotate = 0; // Assume rotation didn't change

            teScaleSlider.value = val;

            teDragActive = true;
            teTransformAppliedInDrag = false;
            handleLiveTransform();

            // End session immediately
            teDragActive = false;
            teTransformAppliedInDrag = false;
        });
    }
    if (teRotateSlider && teRotateInput) {
        teRotateSlider.addEventListener('mousedown', handleDragStart);
        // teRotateSlider.addEventListener('change', () => resetTransformControls(true)); // REMOVED

        teRotateSlider.addEventListener('input', () => {
            teRotateInput.value = teRotateSlider.value;
            handleLiveTransform();
        });
        teRotateInput.addEventListener('change', () => {
            let val = parseFloat(teRotateInput.value);
            if (isNaN(val)) return;

            // When user types a value in the rotation field, they want ABSOLUTE rotation
            // (e.g., typing 0 should reset to 0°, not "rotate by 0°")
            // Use rotateSelection which calls rotateSelectionAbsolute
            console.log('[TRANSFORM] Rotation input changed to ' + val + '°, applying absolute rotation');
            rotateSelection(val);

            // Update slider to match
            teRotateSlider.value = val;

            // Reset transform state
            teDragStartRotate = val;
            teTransformAppliedInDrag = false;
        });
    }

    reloadBtn.addEventListener('click', () => window.location.reload());

    // Add keyboard shortcut for reloading (F5)
    window.addEventListener('keydown', (e) => {
        if (e.key === 'F5') {
            window.location.reload();
        }
        // ESC to close debug log
        if (e.key === 'Escape') {
            const modal = document.getElementById('debug-log-modal');
            if (modal && modal.style.display === 'block') {
                modal.style.display = 'none';
            }
        }
    });

    async function init() {
        // Log immediately using raw csInterface (not Promise wrapper)
        csInterface.evalScript('MDUX_debugLog("[INIT] init() starting...")', function() {});

        try {
            // Re-fetch elements to ensure they exist (in case script ran before DOM)
            teScaleInput = document.getElementById('te-scale');
            teRotateInput = document.getElementById('te-rotate');
            teScaleSlider = document.getElementById('te-scale-slider');
            teRotateSlider = document.getElementById('te-rotate-slider');
            transformEachBtn = document.getElementById('transform-each-btn');
            teResetOriginalBtn = document.getElementById('te-reset-original-btn');
            teLiveOption = document.getElementById('te-live-option');

            csInterface.evalScript('MDUX_debugLog("[INIT] Elements fetched")', function() {});

            initCollapsibleSections();
            csInterface.evalScript('MDUX_debugLog("[INIT] Collapsible sections initialized")', function() {});

            attachListeners();
            csInterface.evalScript('MDUX_debugLog("[INIT] Listeners attached")', function() {});

            // Attach isolation and export listeners
            if (isolatePartsBtn) isolatePartsBtn.addEventListener('click', () => isolate('parts'));
            if (isolateLinesBtn) isolateLinesBtn.addEventListener('click', () => isolate('lines'));
            if (unlockDuctworkBtn) unlockDuctworkBtn.addEventListener('click', () => isolate('unlock'));
            if (createLayersBtn) createLayersBtn.addEventListener('click', () => isolate('create'));
            if (importStylesBtn) importStylesBtn.addEventListener('click', importGraphicStyles);
            if (exportDuctworkBtn) exportDuctworkBtn.addEventListener('click', () => handleExport('ductwork'));
            if (reexportFloorplanBtn) reexportFloorplanBtn.addEventListener('click', () => handleExport('floorplan'));

            csInterface.evalScript('MDUX_debugLog("[INIT] Isolation listeners attached")', function() {});

            // Fix Selection Transform listeners
            if (transformEachBtn) transformEachBtn.addEventListener('click', handleTransformEach);
            if (teResetOriginalBtn) teResetOriginalBtn.addEventListener('click', handleResetOriginal);

            // Reset Strokes and Normalize Ductwork Parts buttons
            if (resetStrokesBtn) {
                resetStrokesBtn.addEventListener('click', async () => {
                    try {
                        await ensureBridgeLoaded();
                        const result = await evalScript('MDUX_resetStrokes()');
                        if (selectionStatus) selectionStatus.textContent = result || 'Strokes reset';
                    } catch (e) {
                        if (selectionStatus) selectionStatus.textContent = 'Error: ' + e.message;
                    }
                });
            }
            if (resetPartsScaleBtn) {
                resetPartsScaleBtn.addEventListener('click', async () => {
                    try {
                        await ensureBridgeLoaded();
                        const result = await evalScript('MDUX_resetDuctworkPartsScale()');
                        if (selectionStatus) selectionStatus.textContent = result || 'Parts scale reset';
                    } catch (e) {
                        if (selectionStatus) selectionStatus.textContent = 'Error: ' + e.message;
                    }
                });
            }

            // Collapsible Document Scale Anchor section
            if (docScaleToggle) {
                docScaleToggle.addEventListener('click', () => {
                    const content = docScaleSection.querySelector('.collapsible-content');
                    const titleSpan = docScaleToggle.querySelector('span');
                    if (docScaleSection.classList.contains('collapsed')) {
                        docScaleSection.classList.remove('collapsed');
                        content.style.display = 'block';
                        titleSpan.textContent = '▼ Document Scale Anchor';
                    } else {
                        docScaleSection.classList.add('collapsed');
                        content.style.display = 'none';
                        titleSpan.textContent = '▶ Document Scale Anchor';
                    }
                });
            }

            // Debug buttons
            const testNoteBtn = document.getElementById('test-note-btn');
            const viewLogBtn = document.getElementById('view-log-btn');
            const clearLogBtn = document.getElementById('clear-log-btn');
            const debugLogModal = document.getElementById('debug-log-modal');
            const closeLogBtn = document.getElementById('close-log-btn');
            const debugLogContent = document.getElementById('debug-log-content');

            if (testNoteBtn) {
                testNoteBtn.addEventListener('click', async () => {
                    try {
                        await ensureBridgeLoaded();
                        const result = await evalScript('MDUX_testNoteProperty()');
                        alert(result);
                    } catch (e) {
                        alert('Error: ' + e.message);
                    }
                });
            }

            if (viewLogBtn) {
                viewLogBtn.addEventListener('click', async () => {
                    try {
                        await ensureBridgeLoaded();
                        const log = await evalScript('MDUX_getDebugLog()');
                        debugLogContent.value = log;
                        debugLogModal.style.display = 'block';
                    } catch (e) {
                        alert('Error: ' + e.message);
                    }
                });
            }

            if (clearLogBtn) {
                clearLogBtn.addEventListener('click', async () => {
                    try {
                        await ensureBridgeLoaded();
                        await evalScript('MDUX_clearDebugLog()');
                        debugStatus.textContent = 'Debug log cleared';
                    } catch (e) {
                        alert('Error: ' + e.message);
                    }
                });
            }

            if (closeLogBtn) {
                closeLogBtn.addEventListener('click', () => {
                    debugLogModal.style.display = 'none';
                });
            }

            csInterface.evalScript('MDUX_debugLog("[INIT] Debug buttons attached")', function() {});

            if (debugStatus) debugStatus.textContent = 'Remote debugging available at http://localhost:8088';
            if (skipOrthoOption) {
                skipOrthoOption.indeterminate = false;
                skipOrthoOption.checked = false;
            }
            if (rotationInput) {
                rotationInput.value = '';
                rotationInput.dataset.autoValue = '';
                rotationInput.dataset.multi = 'false';
            }
            if (skipAllBranchesOption) skipAllBranchesOption.checked = false;
            if (skipFinalOption) skipFinalOption.checked = true;  // Default to checked
            if (createRegisterWiresOption) createRegisterWiresOption.checked = false;
            // Scale controls are hidden - only set values if they exist
            if (scaleSlider) scaleSlider.value = 100;
            if (scaleLabel) scaleLabel.textContent = '100%';
            if (scaleInput) scaleInput.value = '';

            csInterface.evalScript('MDUX_debugLog("[INIT] State initialized")', function() {});

            // Initialize Transform Each controls
            resetTransformControls();

            setProcessStatus('');
            setRevertStatus('');
            setSelectionStatus('');
            setLayerStatus('');
            setImportStatus('');

            csInterface.evalScript('MDUX_debugLog("[INIT] About to load bridge...")', function() {});

            try {
                await ensureBridgeLoaded();
                csInterface.evalScript('MDUX_debugLog("[INIT] Bridge loaded successfully")', function() {});
            } catch (e) {
                var bridgeErrMsg = String(e && e.message ? e.message : e).replace(/'/g, "\\'").replace(/\n/g, " ");
                if (debugStatus) debugStatus.textContent = 'Bridge load failed: ' + bridgeErrMsg;
                csInterface.evalScript('MDUX_debugLog("[INIT] Bridge load failed: ' + bridgeErrMsg + '")', function() {});
            }

            // Test function to verify event is firing
            function testSelectionEvent() {
                evalScript('MDUX_debugLog("[JS] *** afterSelectionChanged EVENT FIRED ***")');
                scheduleSkipOrthoRefresh();
            }

            csInterface.addEventListener('afterSelectionChanged', testSelectionEvent);
            csInterface.addEventListener('documentAfterActivate', scheduleSkipOrthoRefresh);
            csInterface.addEventListener('documentAfterActivate', () => resetTransformControls(true));
            csInterface.addEventListener('documentChanged', scheduleSkipOrthoRefresh);

            csInterface.evalScript('MDUX_debugLog("[INIT] Event listeners registered")', function() {});

            refreshSkipOrthoState().catch(function () { });
            refreshRotationOverrideState().catch(function () { });
            refreshDocScale().catch(function () { });

            csInterface.evalScript('MDUX_debugLog("[INIT] About to create polling interval...")', function() {});

            // Poll every 2 seconds to update selection transform state and rotation override
            const pollInterval = setInterval(function() {
                csInterface.evalScript('MDUX_debugLog("[POLL] tick")', function() {
                    refreshSelectionTransformState().catch(function() {});
                    refreshRotationOverrideState().catch(function() {});
                });
            }, 2000);

            csInterface.evalScript('MDUX_debugLog("[INIT] Polling interval created, ID=' + pollInterval + '")', function() {});

            // Also refresh when panel gets focus
            window.addEventListener('focus', function() {
                csInterface.evalScript('MDUX_debugLog("[FOCUS] Panel got focus")', function() {
                    refreshSelectionTransformState().catch(function() {});
                });
            });

            csInterface.evalScript('MDUX_debugLog("[INIT] Init complete!")', function() {});
        } catch (initError) {
            // Escape single quotes in error message to avoid breaking the evalScript string
            var errMsg = String(initError && initError.message ? initError.message : initError);
            errMsg = errMsg.replace(/'/g, "\\'").replace(/\n/g, " ");
            csInterface.evalScript('MDUX_debugLog("[INIT] ERROR: ' + errMsg + '")', function() {});
        }
    }

    window.addEventListener('beforeunload', function () {
        csInterface.removeEventListener('afterSelectionChanged', scheduleSkipOrthoRefresh);
        csInterface.removeEventListener('documentAfterActivate', scheduleSkipOrthoRefresh);
        csInterface.removeEventListener('documentChanged', scheduleSkipOrthoRefresh);
        evalScript('MDUX_cleanupBridge()');
    });

    document.addEventListener('DOMContentLoaded', init);
})();
