(function () {
    'use strict';

    const csInterface = new CSInterface();
    const processBtn = document.getElementById('process-btn');
    const processStatus = document.getElementById('process-status');
    const toggleIgnoreBtn = document.getElementById('toggle-ignore-btn');
    const applyIgnoreBtn = document.getElementById('apply-ignore-btn');
    const ignoreStatus = document.getElementById('ignore-status');
    const reloadBtn = document.getElementById('reload-btn');
    const debugStatus = document.getElementById('debug-status');

    let ignoreStatusTimer = null;

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
        ignoreStatus.textContent = message || '';
        ignoreStatus.classList.toggle('error', !!isError);
    }

    function normaliseResult(value) {
        if (!value) return { ok: true, value: '' };
        if (typeof value === 'string' && value.indexOf('ERROR:') === 0) {
            return { ok: false, value: value.substring(6) };
        }
        return { ok: true, value: value };
    }

    async function handleProcessClick() {
        processBtn.disabled = true;
        setProcessStatus('Running ductwork script…');
        const result = normaliseResult(await evalScript('MDUX_runMagicDuctwork()'));
        if (result.ok) {
            setProcessStatus('Magic Ductwork completed.');
        } else {
            setProcessStatus('Error: ' + result.value, true);
        }
        processBtn.disabled = false;
    }

    async function toggleIgnoreMode() {
        toggleIgnoreBtn.disabled = true;
        const result = normaliseResult(await evalScript('MDUX_toggleIgnoreModeBridge()'));
        if (!result.ok) {
            setIgnoreStatus('Error: ' + result.value, true);
            toggleIgnoreBtn.disabled = false;
            return;
        }
        await updateIgnoreControls();
        toggleIgnoreBtn.disabled = false;
    }

    async function applyIgnore() {
        applyIgnoreBtn.disabled = true;
        const result = normaliseResult(await evalScript('MDUX_applyIgnoreBridge()'));
        if (!result.ok) {
            setIgnoreStatus('Error: ' + result.value, true);
        } else {
            switch (result.value) {
                case 'IGNORE_ADDED':
                    setIgnoreStatus('Ignore marker added to the selected register endpoint.');
                    break;
                case 'NO_SELECTION':
                    setIgnoreStatus('Select an open duct line to apply an ignore marker.', true);
                    break;
                case 'NO_DOCUMENT':
                    setIgnoreStatus('No Illustrator document is open.', true);
                    break;
                default:
                    setIgnoreStatus(result.value || 'Ignore marker updated.');
                    break;
            }
        }
        applyIgnoreBtn.disabled = false;
        await updateIgnoreControls();
    }

    async function updateIgnoreControls() {
        const state = normaliseResult(await evalScript('MDUX_ignoreModeStatusBridge()'));
        if (!state.ok) {
            setIgnoreStatus('Error: ' + state.value, true);
            toggleIgnoreBtn.textContent = 'Enable Ignore Mode';
            applyIgnoreBtn.disabled = true;
            return;
        }
        const value = state.value || 'inactive';
        if (value.indexOf('active') === 0) {
            toggleIgnoreBtn.textContent = 'Disable Ignore Mode';
            toggleIgnoreBtn.classList.add('active');
            if (value === 'active:no-selection') {
                applyIgnoreBtn.disabled = true;
                setIgnoreStatus('Hover or select an open duct line to preview the register end.');
            } else {
                applyIgnoreBtn.disabled = false;
                const label = value.split(':')[1] || 'selection';
                setIgnoreStatus('Ready to apply ignore marker for "' + label + '".');
            }
            beginIgnoreStatusPolling();
        } else {
            toggleIgnoreBtn.textContent = 'Enable Ignore Mode';
            toggleIgnoreBtn.classList.remove('active');
            applyIgnoreBtn.disabled = true;
            setIgnoreStatus('Ignore mode inactive.');
            endIgnoreStatusPolling();
        }
    }

    function beginIgnoreStatusPolling() {
        if (ignoreStatusTimer) return;
        ignoreStatusTimer = setInterval(updateIgnoreControls, 1500);
    }

    function endIgnoreStatusPolling() {
        if (!ignoreStatusTimer) return;
        clearInterval(ignoreStatusTimer);
        ignoreStatusTimer = null;
    }

    function attachListeners() {
        processBtn.addEventListener('click', handleProcessClick);
        toggleIgnoreBtn.addEventListener('click', toggleIgnoreMode);
        applyIgnoreBtn.addEventListener('click', applyIgnore);
        reloadBtn.addEventListener('click', () => window.location.reload());
        window.addEventListener('beforeunload', () => {
            evalScript('MDUX_cleanupBridge()');
        });
    }

    async function init() {
        attachListeners();
        debugStatus.textContent = 'Remote debugging available at http://localhost:8088';
        await updateIgnoreControls();
    }

    document.addEventListener('DOMContentLoaded', init);
})();
