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
    const clearRotationBtn = document.getElementById('clear-rotation-btn');
    const skipAllBranchesOption = document.getElementById('skip-all-branches-option');
    const skipFinalOption = document.getElementById('skip-final-option');
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
        processStatus.textContent = message || '';
        processStatus.classList.toggle('error', !!isError);
    }

    function setIgnoreStatus(message, isError) {
        // ignoreStatus element removed - function disabled
        // ignoreStatus.textContent = message || '';
        // ignoreStatus.classList.toggle('error', !!isError);
    }

    function setRevertStatus(message, isError) {
        revertStatus.textContent = message || '';
        revertStatus.classList.toggle('error', !!isError);
    }

    function setClearRotationMetadataStatus(message, isError) {
        clearRotationMetadataStatus.textContent = message || '';
        clearRotationMetadataStatus.classList.toggle('error', !!isError);
    }

    function setSelectionStatus(message, isError) {
        selectionStatus.textContent = message || '';
        selectionStatus.classList.toggle('error', !!isError);
    }

    function setLayerStatus(message, isError) {
        layerStatus.textContent = message || '';
        layerStatus.classList.toggle('error', !!isError);
    }

    function setImportStatus(message, isError) {
        importStatus.textContent = message || '';
        importStatus.classList.toggle('error', !!isError);
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
        if (bridgeReloaded) {
            debugStatus.textContent = 'Bridge ready (cached)';
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
        if (skipOrthoRefreshTimer) clearTimeout(skipOrthoRefreshTimer);
        skipOrthoRefreshTimer = setTimeout(function () {
            refreshSkipOrthoState().catch(function () {});
            refreshRotationOverrideState().catch(function () {});
        }, 200);
    }

    async function refreshRotationOverrideState() {
        if (!rotationInput) return;
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
        var prevAuto = rotationInput.dataset.autoValue || '';
        var isFocused = document.activeElement === rotationInput;
        if (!isFocused || formatted !== prevAuto) {
            rotationInput.value = formatted;
            rotationInput.dataset.autoValue = formatted;
        } else {
            rotationInput.dataset.autoValue = formatted;
        }

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
        const isAutoMulti = rotationInput.dataset.multi === 'true' && rotationText === autoValue && rotationText.length > 0;
        if (rotationText && !isAutoMulti) {
            rotationValue = parseFloat(rotationText);
            if (!isFinite(rotationValue)) {
                setProcessStatus('Rotation override must be a valid number.', true);
                processBtn.disabled = false;
                revertBtn.disabled = false;
                return;
            }
        }

        const options = {
            action: 'process',
            skipAllBranchSegments: !!skipAllBranchesOption.checked,
            skipFinalRegisterSegment: !!skipFinalOption.checked
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

    function attachListeners() {
        processBtn.addEventListener('click', handleProcessClick);
        processEmoryBtn.addEventListener('click', handleProcessEmoryClick);
        revertBtn.addEventListener('click', handleRevertPreOrtho);
        clearRotationMetadataBtn.addEventListener('click', handleClearRotationMetadata);
        // applyIgnoreBtn.addEventListener('click', applyIgnore); // Removed - button no longer exists
        clearRotationBtn.addEventListener('click', () => {
            rotationInput.value = '';
            rotationInput.dataset.autoValue = '';
            rotationInput.dataset.multi = 'false';
            rotationInput.placeholder = 'Leave blank to skip';
            setProcessStatus('');
        });
        rotationInput.addEventListener('input', () => {
            rotationInput.dataset.autoValue = '';
            rotationInput.dataset.multi = 'false';
        });
        skipOrthoOption.addEventListener('change', () => {
            skipOrthoOption.indeterminate = false;
        });
        // Mutual exclusivity: only one skip option can be checked at a time
        skipAllBranchesOption.addEventListener('change', () => {
            if (skipAllBranchesOption.checked) {
                skipFinalOption.checked = false;
            }
        });
        skipFinalOption.addEventListener('change', () => {
            if (skipFinalOption.checked) {
                skipAllBranchesOption.checked = false;
            }
        });
        rotate90Btn.addEventListener('click', () => rotateSelection(90));
        rotate45Btn.addEventListener('click', () => rotateSelection(45));
        rotate180Btn.addEventListener('click', () => rotateSelection(180));
        // Custom rotation and scale controls are hidden - only attach listeners if they exist
        if (rotateCustomBtn) {
            rotateCustomBtn.addEventListener('click', () => {
                const value = parseFloat(customRotationInput.value);
                if (!isFinite(value)) {
                    setSelectionStatus('Enter a numeric rotation value.', true);
                    return;
                }
                rotateSelection(value);
            });
        }
        if (scaleSlider) {
            scaleSlider.addEventListener('input', onScaleInput);
        }
        if (applyScaleBtn) {
            applyScaleBtn.addEventListener('click', () => {
                const value = parseFloat(scaleInput.value);
                if (!isFinite(value) || value <= 0) {
                    setSelectionStatus('Enter a valid scale percentage (e.g., 50, 100, 150).', true);
                    return;
                }
                if (value < 25 || value > 500) {
                    setSelectionStatus('Scale must be between 25% and 500%.', true);
                    return;
                }
                applyScale(value);
            });
        }
        if (resetScaleBtn) {
            resetScaleBtn.addEventListener('click', () => {
                scaleSlider.value = 100;
                scaleLabel.textContent = '100%';
                scaleInput.value = '';
                resetScale();
            });
        }
        isolatePartsBtn.addEventListener('click', () => isolate('parts'));
        isolateLinesBtn.addEventListener('click', () => isolate('lines'));
        unlockDuctworkBtn.addEventListener('click', () => isolate('unlock'));
        createLayersBtn.addEventListener('click', () => isolate('create'));
        importStylesBtn.addEventListener('click', importGraphicStyles);
        reloadBtn.addEventListener('click', () => window.location.reload());
    }

    async function init() {
        attachListeners();
        debugStatus.textContent = 'Remote debugging available at http://localhost:8088';
        skipOrthoOption.indeterminate = false;
        skipOrthoOption.checked = false;
        rotationInput.value = '';
        rotationInput.dataset.autoValue = '';
        rotationInput.dataset.multi = 'false';
        skipAllBranchesOption.checked = false;
        skipFinalOption.checked = false;
        createRegisterWiresOption.checked = false;
        // Scale controls are hidden - only set values if they exist
        if (scaleSlider) scaleSlider.value = 100;
        if (scaleLabel) scaleLabel.textContent = '100%';
        if (scaleInput) scaleInput.value = '';
        setProcessStatus('');
        setRevertStatus('');
        setSelectionStatus('');
        setLayerStatus('');
        setImportStatus('');
        // setIgnoreStatus(''); // Removed - status element no longer exists
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            debugStatus.textContent = 'Bridge load failed: ' + (e && e.message ? e.message : e);
        }
        csInterface.addEventListener('afterSelectionChanged', scheduleSkipOrthoRefresh);
        csInterface.addEventListener('documentAfterActivate', scheduleSkipOrthoRefresh);
        csInterface.addEventListener('documentChanged', scheduleSkipOrthoRefresh);
        refreshSkipOrthoState().catch(function () {});
        refreshRotationOverrideState().catch(function () {});
    }

    window.addEventListener('beforeunload', function () {
        csInterface.removeEventListener('afterSelectionChanged', scheduleSkipOrthoRefresh);
        csInterface.removeEventListener('documentAfterActivate', scheduleSkipOrthoRefresh);
        csInterface.removeEventListener('documentChanged', scheduleSkipOrthoRefresh);
        evalScript('MDUX_cleanupBridge()');
    });

    document.addEventListener('DOMContentLoaded', init);
})();
