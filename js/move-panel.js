(() => {
    'use strict';

    const csInterface = new CSInterface();
    const moveToUnitsBtn = document.getElementById('move-to-units-btn');
    const moveToSquareBtn = document.getElementById('move-to-square-btn');
    const moveToRectBtn = document.getElementById('move-to-rect-btn');
    const moveToCircularBtn = document.getElementById('move-to-circular-btn');
    const moveToExhaustBtn = document.getElementById('move-to-exhaust-btn');
    const moveToSecondaryExhaustBtn = document.getElementById('move-to-secondary-exhaust-btn');
    const moveToThermostatsBtn = document.getElementById('move-to-thermostats-btn');
    const moveToIgnoreBtn = document.getElementById('move-to-ignore-btn');
    const moveUseEmoryAssetsOption = document.getElementById('move-use-emory-assets-option');
    const moveStatus = document.getElementById('move-status');
    const debugStatus = document.getElementById('debug-status');
    const reloadBtn = document.getElementById('reload-btn');
    const extensionId = 'com.chris.magicductwork.movepanel';
    const PROCESS_MODE_STORAGE_KEY = 'mdux-process-mode';
    const PROCESS_MODE_EMORY = 'emory';
    const MOVE_PANEL_EMORY_ASSETS_KEY = 'mdux-move-use-emory-assets';
    function reloadExtensionView() {
        try {
            const base = window.location.href.split('?')[0];
            window.location.href = base + '?v=' + Date.now();
            return;
        } catch (e) {
            console.error('View reload failed:', e);
        }
        window.location.reload();
    }

    function readCurrentProcessMode() {
        try {
            return window.localStorage.getItem(PROCESS_MODE_STORAGE_KEY) || '';
        } catch (e) {
            return '';
        }
    }

    function readMovePanelEmoryAssetsPreference() {
        try {
            const stored = window.localStorage.getItem(MOVE_PANEL_EMORY_ASSETS_KEY);
            if (stored === '1') return true;
            if (stored === '0') return false;
        } catch (e) {}
        return readCurrentProcessMode() === PROCESS_MODE_EMORY;
    }

    function writeMovePanelEmoryAssetsPreference(enabled) {
        try {
            window.localStorage.setItem(MOVE_PANEL_EMORY_ASSETS_KEY, enabled ? '1' : '0');
        } catch (e) {}
    }

    function useEmoryAssetsEnabled() {
        return !!(moveUseEmoryAssetsOption && moveUseEmoryAssetsOption.checked);
    }

    function resolveMoveAsset(layerName, defaultFileBaseName) {
        if (!useEmoryAssetsEnabled()) {
            return defaultFileBaseName;
        }
        if (layerName === 'Units') {
            return 'Unit Emory.ai';
        }
        if (layerName === 'Rectangular Registers') {
            return 'Rectangular Register Emory.ai';
        }
        return defaultFileBaseName;
    }

    let bridgeReloaded = false;

    function setMoveStatus(message, isError) {
        moveStatus.textContent = message || '';
        moveStatus.classList.toggle('error', !!isError);
    }

    function normaliseResult(value) {
        if (!value) return { ok: true, value: '' };
        if (typeof value === 'string' && value.indexOf('ERROR:') === 0) {
            return { ok: false, value: value.substring(6) };
        }
        return { ok: true, value: value };
    }

    function evalScript(script) {
        return new Promise((resolve, reject) => {
            csInterface.evalScript(script, result => {
                if (result === 'EvalScript error.') {
                    reject(new Error(result));
                } else {
                    resolve(result);
                }
            });
        });
    }

    async function ensureBridgeLoaded() {
        if (bridgeReloaded) return;
        const scriptPath = csInterface.getSystemPath(CSInterface.SystemPath.EXTENSION) + '/jsx/panel-bridge.jsx';
        await evalScript('$.evalFile("' + scriptPath.replace(/\\/g, '\\\\') + '")');
        bridgeReloaded = true;
    }

    async function moveToLayer(layerName, fileBaseName) {
        console.log('[MOVE] Button clicked - Layer:', layerName, 'File:', fileBaseName);
        setMoveStatus('Moving selection to ' + layerName + '…');

        try {
            console.log('[MOVE] Loading bridge...');
            await ensureBridgeLoaded();
            console.log('[MOVE] Bridge loaded successfully');
        } catch (e) {
            console.error('[MOVE] Bridge load failed:', e);
            setMoveStatus('Bridge load failed: ' + (e && e.message ? e.message : e), true);
            return;
        }

        const payload = JSON.stringify({
            layerName: layerName,
            fileBaseName: fileBaseName
        }).replace(/\\/g, '\\\\').replace(/'/g, "\\'");

        console.log('[MOVE] Payload:', payload);
        console.log('[MOVE] Calling MDUX_moveToLayerBridge...');

        const result = normaliseResult(await evalScript("MDUX_moveToLayerBridge('" + payload + "')"));
        console.log('[MOVE] Result received:', result);

        if (!result.ok) {
            console.error('[MOVE] Error from bridge:', result.value);
            setMoveStatus('Error: ' + result.value, true);
            debugStatus.textContent = 'Move to layer failed: ' + result.value;
            return;
        }

        console.log('[MOVE] Parsing result value:', result.value);
        let stats = null;
        try {
            stats = result.value ? JSON.parse(result.value) : null;
            console.log('[MOVE] Parsed stats:', stats);
        } catch (e) {
            console.error('[MOVE] Failed to parse stats:', e);
            stats = null;
        }

        if (stats && typeof stats.itemsMoved === 'number') {
            console.log('[MOVE] Items moved:', stats.itemsMoved, 'Anchors moved:', stats.anchorsMoved, 'Skipped:', stats.itemsSkipped);
            if (stats.itemsMoved === 0 && stats.anchorsMoved === 0) {
                if (stats.reason === 'no-selection') {
                    setMoveStatus('No items selected.', true);
                } else if (stats.itemsSkipped > 0) {
                    setMoveStatus('Skipped ' + stats.itemsSkipped + ' item(s).', true);
                } else {
                    setMoveStatus('No eligible items found in selection.', true);
                }
            } else {
                const parts = [];
                if (stats.itemsMoved > 0) parts.push(stats.itemsMoved + ' item(s)');
                if (stats.anchorsMoved > 0) parts.push(stats.anchorsMoved + ' anchor(s)');
                let message = 'Moved ' + parts.join(', ') + ' to ' + layerName + '.';
                if (stats.itemsSkipped > 0) {
                    message += ' Skipped ' + stats.itemsSkipped + ' item(s).';
                }
                setMoveStatus(message);
                debugStatus.textContent = 'Moved ' + stats.itemsMoved + ' items, ' + stats.anchorsMoved + ' anchors, skipped ' + stats.itemsSkipped;
            }
        } else {
            console.log('[MOVE] Using fallback status message');
            setMoveStatus(result.value || 'Selection moved to ' + layerName + '.');
        }
    }

    function attachEventListeners() {
        moveToUnitsBtn.addEventListener('click', () => {
            moveToLayer('Units', resolveMoveAsset('Units', 'Unit.ai'));
        });
        moveToSquareBtn.addEventListener('click', () => {
            moveToLayer('Square Registers', 'Square Register.ai');
        });
        moveToRectBtn.addEventListener('click', () => {
            moveToLayer('Rectangular Registers', resolveMoveAsset('Rectangular Registers', 'Rectangular Register.ai'));
        });
        moveToCircularBtn.addEventListener('click', () => {
            moveToLayer('Circular Registers', 'Circular Register.ai');
        });
        moveToExhaustBtn.addEventListener('click', () => {
            moveToLayer('Exhaust Registers', 'Exhaust Register.ai');
        });
        moveToSecondaryExhaustBtn.addEventListener('click', () => {
            moveToLayer('Secondary Exhaust Registers', 'Secondary Exhaust Register.ai');
        });
        moveToThermostatsBtn.addEventListener('click', () => {
            moveToLayer('Thermostats', 'Thermostat.ai');
        });
        moveToIgnoreBtn.addEventListener('click', () => {
            moveToLayer('Ignored', null);
        });
        if (moveUseEmoryAssetsOption) {
            moveUseEmoryAssetsOption.addEventListener('change', () => {
                writeMovePanelEmoryAssetsPreference(!!moveUseEmoryAssetsOption.checked);
            });
        }
        reloadBtn.addEventListener('click', () => {
            reloadExtensionView();
        });
    }

    async function initialise() {
        setMoveStatus('');
        if (moveUseEmoryAssetsOption) {
            moveUseEmoryAssetsOption.checked = readMovePanelEmoryAssetsPreference();
        }
        try {
            await ensureBridgeLoaded();
        } catch (e) {
            console.error('Failed to load bridge on init:', e);
        }
    }

    attachEventListeners();
    initialise();
})();
