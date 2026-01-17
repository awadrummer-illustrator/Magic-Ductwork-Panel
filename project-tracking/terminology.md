# Magic Ductwork Panel - Terminology

This document defines terms used in discussions about this project for quick reference.

## Core Concepts

### Connection Types (Connection Graph Edges)
Connections group paths into runs (connected components). The graph includes an edge between two paths if ANY of the following are true:

1) **Endpoint-to-Endpoint**
Two paths connect if any endpoint pair is within `CONNECTION_DIST` (15pt), unless excluded (ignored anchors, carve-gap separation).

2) **T-Junction (Endpoint-to-Segment)**
A path endpoint connects to the middle of another path segment if:
- distance from endpoint to the segment is <= `T_TOLERANCE` (3pt)
- projection point lies within the segment (0 <= t <= 1)

3) **Segment Intersection with Vertex Requirement**
Two paths connect where they intersect only if:
- the segments intersect geometrically, AND
- at least one path has an anchor at (or extremely near) the intersection (about 0.5pt tolerance)
Crossovers where paths simply cross like an "X" with no vertex at the crossing are NOT connected.

### Ignore Marker
An internal anchor point on a ductwork path that is within 6.5pt of an endpoint. When detected during cleanup, this marker triggers the creation of an ignore part anchor at the endpoint position. The ignore marker itself remains on the ductwork path and creates a very short final segment.

**Behavior**:
- During orthogonalization, the short segment created by the ignore marker is aligned collinear to the adjacent segment and preserves the original marker distance
- After orthogonalization, BOTH the endpoint position and the ignore marker position are added as ignore anchors on the Ignored layer (so neither receives ductwork parts)

### Ignore Gap Part (Auto-Created Gap Marker)
A single-point path automatically created on the "Ignored" layer at carve-out gap endpoints. Marks the exact position where a segment was removed to create a gap at a crossover.

**Purpose**: Block component placement at gap endpoints where paths were split
**Does NOT Block**: Path connections - split halves still connect normally
**Auto-Created At**:
- Each endpoint of a carve-out gap (both `cutBefore` and `cutAfter` positions)
- Always exactly 4.25pt from the original intersection point

**Example**: When two duct runs cross without a connection marker, an 8.5pt gap is created with ignore gap part markers at both ends [gap_start] and [gap_end].

### Ignore Part (User-Placed Endpoint Marker)
A single-point path manually placed on the "Ignored" layer within 4pt (IGNORED_DIST) of a path endpoint. Marks a position where components should NOT be placed, but does NOT block connections.

**Purpose**: Block component placement (registers, units) at specific endpoints
**Does NOT Block**: Path connections - paths still connect normally
**User-Placed At**:
- Near any endpoint where you want to prevent component placement
- Detected by an Ignored-layer anchor within 4pt (IGNORED_DIST) of the endpoint during cleanup

**Example**: Place near a path endpoint to prevent a register from being placed there, while still allowing the path to connect to other paths.

### Connection Marker (User-Placed Connection Forcer)
A single-point path placed on the "Ignored" layer that is NOT near any endpoint (more than 4pt away). FORCES connections at intersections where paths cross, overriding the default crossover behavior.

**Purpose**: Force connections at segment-to-segment intersections (overrides all other connection checks)
**Also Blocks**: Component placement at the intersection point
**User-Placed At**:
- Near intersections where you want to force a connection
- NOT near endpoints (more than 4pt away from all endpoints)

**Example**: Place near a segment-to-segment intersection to force the paths to connect at that point, even if there's no vertex there. This prevents auto-carve-out and merges the paths into a single compound path.

**Note**: Blocking connections is achieved by NOT placing a connection marker, which triggers the default auto-carve-out behavior for unmarked crossovers.

### Carve-Out / Carve-Out Gap
An 8.5pt gap created in a ductwork path where it crosses another path (crossover). The gap is created by:
1. Splitting the path into two halves
2. Removing the segment between `cutBefore` and `cutAfter` points (each 4.25pt from intersection)
3. Placing ignore markers at both gap endpoints

### Crossover
Where two ductwork paths cross each other WITHOUT a connection (separate duct runs passing over/under each other). Identified by:
- Segment intersection exists
- Neither path has a vertex at the intersection point
- Results in carve-out gap creation

### Connection Marker a.k.a. Intersection Marker
A single-point path (solo anchor) placed near a segment-to-segment intersection to indicate that a connection should be made there. When detected, a vertex is inserted at the intersection point on one of the paths, converting a crossover into a real connection.

### T-Junction
A connection where one path's endpoint meets another path's segment (not at a vertex). Branch lines connecting to trunk lines are T-junctions.

### Unit Anchor
The position where a Unit component sits (or will be placed). A Unit anchor can be detected by:
- A Unit component on the "Units" layer
- A Thermostat Line endpoint within the unit placement tolerance (~10pt) acting as a proxy
Thermostat Lines are a proxy only and are not required for unit detection.

### Thermostat Line Endpoint (Unit Proxy)
An endpoint on the "Thermostat Lines" layer used as a proxy for a Unit location. Thermostat Lines are curved wire drawings and should never be orthogonalized.

### AHU Anchor
The position where an AHU component sits (on the "AHU" layer). The AHU can align to an endpoint or to the middle of a duct segment (loop feed).

### Duct Role Classification (Trunk vs Branch)
Role is determined by connection topology plus seed anchors:
- Blue/Orange Ductwork: endpoints near Unit anchors seed Trunk
- Green/Light Green/Light Orange Ductwork: endpoints near AHU anchors seed Trunk; endpoints near Unit anchors are Branches
- Endpoint-to-endpoint connections keep the same role
- Endpoint-to-segment or segment-to-segment connections create a Branch off the parent line
Registers typically live on Branches for Blue/Orange; a rare trunk register on the far end of a unit-run does NOT change the trunk role.

### Trunk
A main ductwork path that typically:
- Originates from or connects to a Unit anchor (Blue/Orange) or an AHU anchor (Green/Light Green/Light Orange)
- May terminate at a register or continue to other connections (a register does not disqualify a trunk)
- May have one or more branches connecting to it (branches can also branch)
- Forms the "spine" of a duct run that branches distribute airflow from

Trunks are identified by their role in the ductwork topology, not by any special marking. A trunk with branches forms a tree-like structure where the trunk is the main line and branches are offshoots.

### Branch
A ductwork path that connects to another path (typically a trunk) via T-junction:
- One endpoint connects to a trunk's segment (not at an existing vertex)
- Other endpoint typically terminates at a register or another connection
- Creates a T-junction connection point when processed
- Multiple branches can connect to the same trunk at different points

Branches distribute airflow from trunks to individual registers or zones. In the connection logic, branches are detected when a path endpoint qualifies as a T-junction under `T_JUNCTION_DIST` (3pt) endpoint-to-segment detection. Branches can also branch.

### Run
A **run** is a **connected component**: a group of PathItems that are connected (directly or transitively) by the connection detection rules. Each run is merged into **one compound path** during the Compounding phase.

Notes:
- Separate runs remain separate compound paths even if they geometrically cross as a **crossover** (no qualifying connection).
- Runs are produced by grouping paths via the connection graph (union-find / connected components).

### Connected Component
A group of paths where every path can reach every other path through one or more detected connections (graph connectivity). Each connected component becomes one "run" and is compounded into one compound path.

### Connection Detection
The algorithm that builds a connection graph between paths. Connections are added only when they meet one of the defined connection types (endpoint-to-endpoint, T-junction [Endpoint to Segment], intersection [segment to segment]), and only if they do not violate exclusion rules (ignored anchors, carve-gap separation, crossovers without vertices).

### Compound Path
Multiple connected ductwork paths merged into a single compound path item. Paths are grouped into connected components from detected connections, then each component is merged into one compound path during Compounding.

Connection types that create edges in the connection graph:
- **Endpoint-to-endpoint**: endpoints within `CONNECTION_DIST` (15pt)
- **T-junction (endpoint-to-segment)**: endpoint within `T_TOLERANCE` (3pt) of a segment, with the projection falling inside the segment
- **Intersection-with-vertex**: segments intersect geometrically AND at least one path has an anchor at (or extremely near) the intersection. Crossovers without a vertex at the crossing are NOT connections.

### Split Path / Split Halves
When a path is split at a crossover or carve-out point:
- **Original path** becomes one half (keeps start to cut point)
- **Duplicate path** becomes other half (keeps cut point to end)
- Both halves should remain in the same compound (forced connection)

### Deleted Segment
The small segment removed during carve-out. Now stored on hidden/locked "Deleted Segments" layer for potential recovery.

## Layers

| Layer Name | Purpose |
|------------|---------|
| Blue Ductwork | Main supply ductwork paths |
| Green Ductwork | Exhaust/return ductwork paths |
| Light Green Ductwork | Exhaust/return secondary ductwork paths |
| Orange Ductwork | Exhaust/supply ductwork paths |
| Light Orange Ductwork | Exhaust/return secondary ductwork paths |
| Thermostat Lines | Curved wire paths (never orthogonalized) |
| Ignored | Contains ignore anchors (single-point paths marking positions to skip) |
| Deleted Segments | Backup of removed carve-out segments (hidden/locked) |
| Units | Component placement layer |
| AHU | AHU component placement layer |
| Square Registers | Square register components |
| Exhaust Registers | Exhaust register components |
| Thermostats | Thermostat components |

## Processing Phases (Order)

1. **Role Classification (Trunk/Branch)** - Seed trunks from Unit/AHU anchors and propagate roles across connections (used by skip-final-ortho)
2. **Orthogonalization** - Snaps ductwork paths to horizontal/vertical (Thermostat Lines are excluded)
3. **Small Segment Detection** - Finds existing crossover indicators
4. **Marker-Triggered Intersections** - Adds vertices at intersections marked by solo anchors
5. **AUTO-CARVE** - Creates gaps at unmarked crossover intersections
6. **Connection Finding** - Detects path connections (anchor-to-anchor, T-junction [Endpoint to Segment], intersection [segment to segment])
7. **Crossover Filtering** - Removes connections between crossover paths
8. **Compounding** - Merges connected paths into compound paths
9. **Component Placement** - Places registers, units, thermostats

## Key Variables

| Variable | Purpose |
|----------|---------|
| `ignoredAnchors` | Array of [x,y] positions to skip for component placement |
| `CARVE_BLOCKED_CONNECTIONS` | Runtime array of {carvedPath, crossingPath} pairs to block (empty on rerun) |
| `crossoverInfo` | Array of detected crossover info objects |
| `splitPathPairs` | Array of {pathA, pathB} for split path halves that should stay connected |

## Spatial Thresholds (CRITICAL - Check these first for connection/compounding bugs!)

| Constant | Value | Location | Purpose |
|----------|-------|----------|---------|
| `T_JUNCTION_DIST` | 3pt | Line ~9924 | Max distance for T-junction detection (endpoint to segment). **Was 25pt - caused unrelated paths to merge!** |
| `AUTO_CARVE_HALF_WIDTH` | 4.25pt | Line ~14300 | Half the carve-out gap width. T-junctions at exactly this distance are skipped as false positives. |
| `MIN_DIST` | 0.5pt | Line ~9923 | Minimum distance - paths closer are likely duplicates |
| `PATH_ANCHOR_TOLERANCE` | 10pt | Line ~9925 | Distance threshold for path vertex at intersection check |
| `ENDPOINT_TOLERANCE` | 15pt | Line ~10172 | Extended anchor-to-anchor connection tolerance |
| `CONNECTION_DIST` | 15pt | Line ~10561 | Endpoint-to-endpoint connection threshold. Same effective value as `ENDPOINT_TOLERANCE`. |
| `T_TOLERANCE` | 3pt | Line ~10492 | T-junction endpoint-to-segment tolerance. Same effective value as `T_JUNCTION_DIST`. |
| `VERTEX_INTERSECTION_TOL` | 0.5pt | Line ~106xx | Vertex-required intersection tolerance for treating an intersection as a connection. |
| `IGNORED_DIST` | 4pt | Line ~330 | Near ignored anchors: block connections or placement when near ignored anchors. |
| `ENDPOINT_INTERNAL_THRESHOLD` | 6.5pt | Line ~15200 | Internal anchor near endpoint threshold for detecting ignore markers. |

**TROUBLESHOOTING:** If unrelated paths are being incorrectly merged into compound paths, CHECK THE SPATIAL THRESHOLDS FIRST. Look at the debug log for "T-junction detected" messages showing the distances - if paths 20+ pt apart are being connected, a threshold is too high.

## Known Issues

1. ~~**Ignore anchors before ortho**: Ignore anchors created at endpoints with ignore markers (close internal anchors) were being placed before orthogonalization, causing position mismatch after ortho.~~ **FIXED** (2026-01-08): Now stores path references during cleanup, reads actual endpoint coordinates AFTER orthogonalization.

2. **Rerun compound separation**: On second run, some split paths may not end up in the correct compound path due to timing/detection issues.
