#target illustrator

#include "./register-ignore.jsx"

if (typeof MDUX === "undefined") {
    var MDUX = {};
}

function MDUX_extensionRoot() {
    var file = File($.fileName);
    return file.parent;
}

function MDUX_runMagicDuctwork() {
    try {
        var root = MDUX_extensionRoot();
        var scriptFile = File(root + "/magic-final.jsx");
        if (!scriptFile.exists) {
            return "ERROR:magic-final.jsx not found";
        }
        $.evalFile(scriptFile);
        return "OK";
    } catch (e) {
        return "ERROR:" + e;
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

function MDUX_applyIgnoreBridge() {
    try {
        return MDUX_applyIgnoreToCurrent();
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
