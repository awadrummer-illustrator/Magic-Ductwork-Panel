# Magic Ductwork Panel - AI Information

## Project Overview

Adobe Illustrator CEP extension for automated ductwork diagram generation with orthogonalization, T-junction detection, and ignore marker support.

Important: Emory Mode launched from this panel routes to the separate `Emory-Ductwork-Panel` native plugin target `EmoryDuctwork.aip`. Bugs in `Process Emory Ductwork` are not necessarily in this repo's `ProcessDuctwork.aip` path.

## Architecture

- **Frontend**: HTML/JavaScript panel UI
- **Backend**: ExtendScript JSX (ES3 syntax)
- **Acceleration**: Python geometry engine for 30-50x faster orthogonalization
- **Debug Logs**: `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\Debug\`
- **Deployment**: `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\`

## Key Files

- `jsx/magic-final.jsx` - Main ductwork generation logic (14,000+ lines)
- `python/geometry_engine.py` - Python-accelerated orthogonalization
- `jsx/python-bridge.jsx` - Bridge between ExtendScript and Python

## Critical Known Issues & Solutions

### Issue: Ignore Marker Distance Not Preserved During Multi-Path Orthogonalization

**Problem Description:**
When selecting multiple paths together for processing, paths with ignore markers (internal anchors within 4pt of endpoints) would have their endpoint snap to the marker position, collapsing the 1.1pt distance segment.

**Root Cause:**
Python's orthogonalization uses a `SNAP_THRESHOLD` of 3pt. When paths with ignore markers were sent to Python, the marker (1.1pt from endpoint) would snap to the endpoint because 1.1pt < 3pt threshold. This happened even when both points were "locked" - locking only prevented movement relative to OTHER paths, not snapping to each other.

**Attempted Fixes That Failed:**
1. Locking just the marker anchor - Python moved the endpoint to marker
2. Locking both marker AND endpoint - Python still snapped them together (within snap threshold)
3. POST-PYTHON-ORTHO adjustment after Python - endpoint already collapsed, wrong direction detected
4. Removing duplicate `restoreEndpointConnections()` - helped but didn't solve root cause

**Correct Solution:**
**EXCLUDE paths with ignore markers from Python processing entirely.**

Implementation in `jsx/magic-final.jsx` around line 13994:
```javascript
// EXCLUDE paths with ignore markers from Python processing entirely
// Python's SNAP_THRESHOLD (3pt) will snap the marker and endpoint together (they're only 1.1pt apart)
var pathsForPython = [];
var excludedPaths = [];
for (var pyPathIdx = 0; pyPathIdx < geometryPaths.length; pyPathIdx++) {
    var hasIgnoreMarker = false;
    for (var checkIdx = 0; checkIdx < ORTHO_IGNORE_MARKER_PATHS.length; checkIdx++) {
        if (ORTHO_IGNORE_MARKER_PATHS[checkIdx].path === geometryPaths[pyPathIdx]) {
            hasIgnoreMarker = true;
            excludedPaths.push(pyPathIdx);
            break;
        }
    }
    if (!hasIgnoreMarker) {
        pathsForPython.push(geometryPaths[pyPathIdx]);
    }
}
// Send only pathsForPython (not geometryPaths) to Python
var pyOrthoResult = PythonBridge.orthogonalize(pathsForPython, SNAP_THRESHOLD, lockedPointsForPython);
```

**Why This Works:**
- Single-path processing: Path with ignore marker processed alone first, gets correct 1.1pt distance at correct angle
- Multi-path processing: Ignore marker path already correct, excluded from Python array
- Python orthogonalizes all OTHER paths (connecting paths, etc.) without touching the ignore marker path
- T-junction restoration still works via POST-ORTHO-TJ for the connecting path endpoint

**Key Lesson:**
When Python's snap threshold conflicts with required distances, exclude those paths from Python processing rather than trying to lock or adjust after the fact.

## Processing Flow

1. **Cleanup Phase**: Detect internal anchors within 4pt of endpoints, store in `ORTHO_IGNORE_MARKER_PATHS`
2. **Pre-Ortho T-Junction Detection**: Find T-junctions before orthogonalization, store marker-to-T distances
3. **Python Orthogonalization**: Exclude ignore marker paths, send only other paths to Python
4. **POST-PYTHON-ORTHO**: Skip adjustment for excluded paths (already correct)
5. **POST-ORTHO-TJ**: Snap connecting path endpoints to segments, skip marker/endpoint adjustment for segments with ignore markers

## Terminology

- **Ignore Marker**: Internal anchor within 4pt of endpoint that triggers ignore part creation
- **T-Junction**: Connection where one path's endpoint meets another path's segment (not at vertex)
- **SNAP_THRESHOLD**: 3pt distance within which Python snaps points together
- **geometryPaths**: All paths being processed
- **ORTHO_IGNORE_MARKER_PATHS**: Array of `{path, endpoint, ignoreMarkerIndex, originalDistance}` for paths with ignore markers
- **Locked Points**: Anchors that Python shouldn't move during orthogonalization

## ExtendScript Constraints

- ES3 syntax only (no let/const, no arrow functions, no template literals)
- Must use `var` for all variables
- PathItem anchors have both `anchor` position and `leftDirection`/`rightDirection` Bezier handles
- Handles must equal anchor position for corner points (straight segments)

## Debug Log Patterns

Key patterns to search for when debugging:
- `[CLEANUP]` - Internal anchor detection and cleanup
- `[PRE-ORTHO-TJ]` - T-junction detection before orthogonalization
- `[PYTHON-ORTHO]` - Python orthogonalization messages
- `[POST-PYTHON-ORTHO]` - Ignore marker endpoint adjustment after Python
- `[POST-ORTHO-TJ]` - T-junction restoration after orthogonalization

## Common Debugging Commands

Read latest log:
```powershell
Get-ChildItem 'C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\Debug\' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { Get-Content $_.FullName }
```

Search for specific patterns:
```powershell
Get-Content <logfile> | Select-String -Pattern "PYTHON-ORTHO|POST-ORTHO"
```

## Deployment

Read the deployment instructions in this repo before claiming a change is deployed. For this project that means checking `README.md`, `DEPLOYMENT_INSTRUCTIONS.md`, and `cpp-plugin/AI Info.md`.

After panel-only changes:
```powershell
Copy-Item 'e:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Magic-Ductwork-Panel\jsx\magic-final.jsx' 'C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\jsx\magic-final.jsx' -Force
```

After plugin/C++ changes:
```powershell
schtasks /run /tn "Reload Illustrator Ductwork"
```

Preferred wrapper:
```powershell
powershell -ExecutionPolicy Bypass -File ".\tools\reload-illustrator-ductwork.ps1"
```

Do not say deployment is complete until the task finished with `Last Result: 0`, the installed `ProcessDuctwork.aip` matches the built plugin, and Illustrator has relaunched.
