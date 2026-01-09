/**
 * Geometry Bridge - Node.js bridge to Python geometry server
 *
 * This module manages a persistent Python HTTP server for fast geometry operations.
 * Eliminates the ~7 second startup overhead per Python call.
 */

const { spawn } = require('child_process');
const http = require('http');
const path = require('path');

const GEOMETRY_SERVER_PORT = 8765;
let serverProcess = null;
let serverReady = false;

/**
 * Get the path to the Python server script
 */
function getServerPath() {
    // Get extension root from CSInterface
    const csInterface = new CSInterface();
    const extRoot = csInterface.getSystemPath(CSInterface.SystemPath.EXTENSION);
    return path.join(extRoot, 'python', 'geometry_server.py');
}

/**
 * Check if the server is running
 */
function checkServerHealth() {
    return new Promise((resolve) => {
        const req = http.request({
            hostname: '127.0.0.1',
            port: GEOMETRY_SERVER_PORT,
            path: '/health',
            method: 'GET',
            timeout: 2000
        }, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                try {
                    const result = JSON.parse(data);
                    resolve(result.status === 'ok');
                } catch {
                    resolve(false);
                }
            });
        });
        req.on('error', () => resolve(false));
        req.on('timeout', () => { req.destroy(); resolve(false); });
        req.end();
    });
}

/**
 * Start the Python geometry server
 */
async function startServer() {
    // Check if already running
    if (await checkServerHealth()) {
        console.log('[GEOM-BRIDGE] Server already running');
        serverReady = true;
        return true;
    }

    const serverPath = getServerPath();
    console.log('[GEOM-BRIDGE] Starting server:', serverPath);

    return new Promise((resolve) => {
        try {
            serverProcess = spawn('python', [serverPath, '--port', GEOMETRY_SERVER_PORT.toString()], {
                stdio: ['ignore', 'pipe', 'pipe'],
                windowsHide: true
            });

            let startupTimeout = setTimeout(() => {
                console.log('[GEOM-BRIDGE] Server startup timeout');
                resolve(false);
            }, 15000);

            serverProcess.stdout.on('data', (data) => {
                const msg = data.toString().trim();
                console.log('[GEOM-SERVER]', msg);
                if (msg.startsWith('READY:')) {
                    clearTimeout(startupTimeout);
                    serverReady = true;
                    console.log('[GEOM-BRIDGE] Server ready on port', GEOMETRY_SERVER_PORT);
                    resolve(true);
                } else if (msg.startsWith('ALREADY_RUNNING:')) {
                    clearTimeout(startupTimeout);
                    serverReady = true;
                    resolve(true);
                } else if (msg.startsWith('ERROR:')) {
                    clearTimeout(startupTimeout);
                    resolve(false);
                }
            });

            serverProcess.stderr.on('data', (data) => {
                console.log('[GEOM-SERVER-ERR]', data.toString().trim());
            });

            serverProcess.on('error', (err) => {
                console.error('[GEOM-BRIDGE] Failed to start server:', err);
                clearTimeout(startupTimeout);
                resolve(false);
            });

            serverProcess.on('exit', (code) => {
                console.log('[GEOM-BRIDGE] Server exited with code', code);
                serverReady = false;
                serverProcess = null;
            });

        } catch (err) {
            console.error('[GEOM-BRIDGE] Error starting server:', err);
            resolve(false);
        }
    });
}

/**
 * Stop the server
 */
async function stopServer() {
    if (serverProcess) {
        // Try graceful shutdown first
        try {
            await makeRequest('shutdown', {}, 'GET');
        } catch {
            // Force kill if graceful fails
            serverProcess.kill();
        }
        serverProcess = null;
        serverReady = false;
    }
}

/**
 * Make a request to the geometry server
 */
function makeRequest(operation, data, method = 'POST') {
    return new Promise((resolve, reject) => {
        const postData = method === 'POST' ? JSON.stringify(data) : '';

        const options = {
            hostname: '127.0.0.1',
            port: GEOMETRY_SERVER_PORT,
            path: '/' + operation,
            method: method,
            timeout: 120000, // 2 minute timeout for large operations
            headers: method === 'POST' ? {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(postData)
            } : {}
        };

        const req = http.request(options, (res) => {
            let responseData = '';
            res.on('data', chunk => responseData += chunk);
            res.on('end', () => {
                try {
                    const result = JSON.parse(responseData);
                    if (res.statusCode >= 400) {
                        reject(new Error(result.error || 'Server error'));
                    } else {
                        resolve(result);
                    }
                } catch (e) {
                    reject(new Error('Invalid JSON response'));
                }
            });
        });

        req.on('error', reject);
        req.on('timeout', () => {
            req.destroy();
            reject(new Error('Request timeout'));
        });

        if (method === 'POST') {
            req.write(postData);
        }
        req.end();
    });
}

/**
 * Execute a geometry operation
 */
async function executeOperation(operation, paths, params = {}) {
    // Ensure server is running
    if (!serverReady) {
        const started = await startServer();
        if (!started) {
            throw new Error('Failed to start geometry server');
        }
    }

    // Make the request
    const startTime = Date.now();
    const result = await makeRequest(operation, { paths, params });
    const totalTime = Date.now() - startTime;

    console.log(`[GEOM-BRIDGE] ${operation}: ${paths.length} paths in ${totalTime}ms (server: ${result.time_ms?.toFixed(1) || '?'}ms)`);

    return result;
}

// Public API
window.GeometryBridge = {
    startServer,
    stopServer,
    checkServerHealth,

    // Geometry operations
    orthogonalize: (paths, params) => executeOperation('orthogonalize', paths, params),
    findConnections: (paths, params) => executeOperation('find_connections', paths, params),
    buildGroups: (paths, params) => executeOperation('build_groups', paths, params),
    detectIntersections: (paths, params) => executeOperation('detect_intersections', paths, params),
    snapAnchors: (paths, params) => executeOperation('snap_anchors', paths, params),

    // Status
    isReady: () => serverReady,
    getPort: () => GEOMETRY_SERVER_PORT
};

// Auto-start server when module loads
console.log('[GEOM-BRIDGE] Module loaded, starting server...');
startServer().then(success => {
    if (success) {
        console.log('[GEOM-BRIDGE] Server started successfully');
    } else {
        console.warn('[GEOM-BRIDGE] Server failed to start - will use fallback');
    }
});

// Cleanup on window unload
window.addEventListener('beforeunload', () => {
    stopServer();
});
