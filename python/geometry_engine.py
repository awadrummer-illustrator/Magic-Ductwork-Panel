"""
Magic Ductwork Geometry Engine
High-performance geometry operations using Shapely + spatial indexing.
Called from ExtendScript to offload heavy computation.

Usage:
    python geometry_engine.py <operation> < input.json > output.json

Operations:
    find_connections - Find all connected path pairs
    detect_intersections - Find segment intersection points
    classify_junctions - Classify T-junctions vs crossovers
"""

import json
import sys
import time
from typing import List, Dict, Tuple, Any, Optional

# Shapely imports
from shapely.geometry import LineString, Point, MultiPoint
from shapely.strtree import STRtree
from shapely import get_coordinates
import numpy as np

# Configuration
CLOSE_DIST = 10.0  # px for loose connection grouping
CONNECTION_DIST = 2.0  # px stricter threshold for actual compounding
T_JUNCTION_DIST = 3.0  # Tolerance for T-junction detection
ANGLE_THRESHOLD_DEG = 20.0  # Degrees for directional alignment
MIN_DIST = 0.5  # Minimum distance - closer paths are likely duplicates
PATH_ANCHOR_TOLERANCE = 10.0  # Distance threshold for path vertex at intersection


def log(msg: str):
    """Log to stderr so it doesn't pollute JSON output."""
    print(f"[PYGEOM] {msg}", file=sys.stderr)


def paths_to_linestrings(paths_data: List[Dict]) -> List[Optional[LineString]]:
    """Convert path data from ExtendScript to Shapely LineStrings."""
    lines = []
    for p in paths_data:
        points = p.get('points', [])
        if len(points) >= 2:
            coords = [(pt['x'], pt['y']) for pt in points]
            lines.append(LineString(coords))
        elif len(points) == 1:
            # Single-point path (anchor marker)
            lines.append(Point(points[0]['x'], points[0]['y']))
        else:
            lines.append(None)
    return lines


def get_endpoints(line: LineString) -> Tuple[Point, Point]:
    """Get start and end points of a LineString."""
    coords = list(line.coords)
    return Point(coords[0]), Point(coords[-1])


def get_segment_direction(p1: Tuple[float, float], p2: Tuple[float, float]) -> np.ndarray:
    """Get normalized direction vector between two points."""
    vec = np.array([p2[0] - p1[0], p2[1] - p1[1]])
    norm = np.linalg.norm(vec)
    if norm < 1e-10:
        return np.array([0.0, 0.0])
    return vec / norm


def angle_between_vectors(v1: np.ndarray, v2: np.ndarray) -> float:
    """Calculate angle in degrees between two direction vectors."""
    dot = np.clip(np.dot(v1, v2), -1.0, 1.0)
    return np.degrees(np.arccos(abs(dot)))  # abs() to handle opposite directions


def get_adjacent_directions(line: LineString, point_idx: int) -> List[np.ndarray]:
    """Get direction vectors of segments adjacent to a point."""
    coords = list(line.coords)
    dirs = []

    if point_idx > 0:
        dirs.append(get_segment_direction(coords[point_idx - 1], coords[point_idx]))
    if point_idx < len(coords) - 1:
        dirs.append(get_segment_direction(coords[point_idx], coords[point_idx + 1]))

    return dirs


def point_near_vertex(point: Point, line: LineString, tolerance: float) -> bool:
    """Check if a point is near any vertex of a line."""
    for coord in line.coords:
        if point.distance(Point(coord)) <= tolerance:
            return True
    return False


def is_crossover(int_point: Point, line_a: LineString, line_b: LineString, tolerance: float = PATH_ANCHOR_TOLERANCE) -> bool:
    """
    Check if an intersection point is a crossover (no vertex at intersection).
    Returns True if neither line has a vertex at the intersection point.
    """
    a_has_vertex = point_near_vertex(int_point, line_a, tolerance)
    b_has_vertex = point_near_vertex(int_point, line_b, tolerance)
    return not (a_has_vertex or b_has_vertex)


def closest_point_on_segment(seg_start: Tuple, seg_end: Tuple, point: Tuple) -> Tuple[Tuple[float, float], float]:
    """
    Find closest point on a line segment to a given point.
    Returns (closest_point, t_parameter) where t is 0-1 along segment.
    """
    ax, ay = seg_start
    bx, by = seg_end
    px, py = point

    dx = bx - ax
    dy = by - ay
    seg_len_sq = dx * dx + dy * dy

    if seg_len_sq < 1e-10:
        return (ax, ay), 0.0

    t = ((px - ax) * dx + (py - ay) * dy) / seg_len_sq
    t = max(0.0, min(1.0, t))

    closest = (ax + t * dx, ay + t * dy)
    return closest, t


def find_connections(paths_data: List[Dict], max_dist: float = CLOSE_DIST) -> Dict:
    """
    Find all connected path pairs using spatial indexing.
    This is the O(n log n) replacement for the O(n^2) ExtendScript version.
    """
    start_time = time.time()

    lines = paths_to_linestrings(paths_data)
    valid_indices = [i for i, line in enumerate(lines) if line is not None and isinstance(line, LineString)]
    valid_lines = [lines[i] for i in valid_indices]

    if len(valid_lines) < 2:
        return {'connections': [], 'time_ms': 0, 'path_count': len(paths_data)}

    log(f"Processing {len(valid_lines)} valid paths...")

    # Build spatial index - this is what makes it fast
    tree = STRtree(valid_lines)

    connections = []
    ignored_anchors = []  # Points where we shouldn't place components
    checked_pairs = set()

    for idx_a, line_a in enumerate(valid_lines):
        original_idx_a = valid_indices[idx_a]

        # Query spatial index for nearby lines (buffer by max_dist)
        buffered = line_a.buffer(max_dist + T_JUNCTION_DIST)
        candidate_indices = tree.query(buffered)

        for idx_b in candidate_indices:
            if idx_b <= idx_a:
                continue

            original_idx_b = valid_indices[idx_b]
            pair_key = (original_idx_a, original_idx_b)
            if pair_key in checked_pairs:
                continue
            checked_pairs.add(pair_key)

            line_b = valid_lines[idx_b]
            connected = False
            connection_type = None
            connection_point = None

            # Check for duplicate paths (same geometry at same location)
            if len(list(line_a.coords)) == len(list(line_b.coords)):
                coords_a = np.array(line_a.coords)
                coords_b = np.array(line_b.coords)
                if np.allclose(coords_a, coords_b, atol=MIN_DIST):
                    continue  # Skip duplicates

            # 1. Anchor-to-anchor with directional alignment
            coords_a = list(line_a.coords)
            coords_b = list(line_b.coords)

            for ai, pt_a in enumerate(coords_a):
                if connected:
                    break
                for bi, pt_b in enumerate(coords_b):
                    dist = Point(pt_a).distance(Point(pt_b))
                    if MIN_DIST <= dist <= max_dist:
                        # Check directional alignment
                        dirs_a = get_adjacent_directions(line_a, ai)
                        dirs_b = get_adjacent_directions(line_b, bi)

                        aligned = False
                        for da in dirs_a:
                            for db in dirs_b:
                                if angle_between_vectors(da, db) <= ANGLE_THRESHOLD_DEG:
                                    aligned = True
                                    break
                            if aligned:
                                break

                        if aligned:
                            connected = True
                            connection_type = 'anchor_aligned'
                            connection_point = [(pt_a[0] + pt_b[0]) / 2, (pt_a[1] + pt_b[1]) / 2]
                            break

            # 2. T-junction detection (point-to-segment)
            if not connected:
                for ai, pt_a in enumerate(coords_a):
                    if connected:
                        break
                    for bi in range(len(coords_b) - 1):
                        closest, t = closest_point_on_segment(coords_b[bi], coords_b[bi + 1], pt_a)
                        dist = Point(pt_a).distance(Point(closest))

                        if MIN_DIST <= dist <= T_JUNCTION_DIST and 0 < t < 1:
                            # Skip if at carve-gap distance (4.25pt)
                            if abs(dist - 4.25) < 0.5:
                                continue

                            # Check for crossover
                            int_pt = Point(closest)
                            if is_crossover(int_pt, line_a, line_b, T_JUNCTION_DIST):
                                continue  # Crossover, not a real connection

                            connected = True
                            connection_type = 't_junction'
                            connection_point = [closest[0], closest[1]]
                            break

                # Check reverse direction too
                if not connected:
                    for bi, pt_b in enumerate(coords_b):
                        if connected:
                            break
                        for ai in range(len(coords_a) - 1):
                            closest, t = closest_point_on_segment(coords_a[ai], coords_a[ai + 1], pt_b)
                            dist = Point(pt_b).distance(Point(closest))

                            if MIN_DIST <= dist <= T_JUNCTION_DIST and 0 < t < 1:
                                if abs(dist - 4.25) < 0.5:
                                    continue

                                int_pt = Point(closest)
                                if is_crossover(int_pt, line_a, line_b, T_JUNCTION_DIST):
                                    continue

                                connected = True
                                connection_type = 't_junction'
                                connection_point = [closest[0], closest[1]]
                                break

            # 3. Extended endpoint-to-endpoint (for branches)
            if not connected:
                endpoints_a = [coords_a[0], coords_a[-1]]
                endpoints_b = [coords_b[0], coords_b[-1]]

                for ep_a in endpoints_a:
                    if connected:
                        break
                    for ep_b in endpoints_b:
                        dist = Point(ep_a).distance(Point(ep_b))
                        if MIN_DIST <= dist <= 15:  # Extended tolerance for endpoints
                            mid = Point((ep_a[0] + ep_b[0]) / 2, (ep_a[1] + ep_b[1]) / 2)
                            if not is_crossover(mid, line_a, line_b, 15):
                                connected = True
                                connection_type = 'endpoint'
                                connection_point = [mid.x, mid.y]
                                break

            # 4. Segment intersection with vertex check
            if not connected:
                if line_a.intersects(line_b):
                    intersection = line_a.intersection(line_b)
                    if not intersection.is_empty:
                        if hasattr(intersection, 'x'):
                            int_pt = intersection
                        else:
                            int_pt = intersection.centroid

                        if not is_crossover(int_pt, line_a, line_b):
                            connected = True
                            connection_type = 'intersection_vertex'
                            connection_point = [int_pt.x, int_pt.y]
                            ignored_anchors.append([int_pt.x, int_pt.y])

            if connected:
                connections.append({
                    'a': original_idx_a,
                    'b': original_idx_b,
                    'type': connection_type,
                    'point': connection_point
                })

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Found {len(connections)} connections in {elapsed_ms:.1f}ms")

    return {
        'connections': connections,
        'ignored_anchors': ignored_anchors,
        'time_ms': elapsed_ms,
        'path_count': len(paths_data),
        'valid_path_count': len(valid_lines)
    }


def detect_intersections(paths_data: List[Dict]) -> Dict:
    """
    Find all intersection points between path segments.
    Returns detailed intersection info for each pair.
    """
    start_time = time.time()

    lines = paths_to_linestrings(paths_data)
    valid_indices = [i for i, line in enumerate(lines) if line is not None and isinstance(line, LineString)]
    valid_lines = [lines[i] for i in valid_indices]

    if len(valid_lines) < 2:
        return {'intersections': [], 'time_ms': 0}

    tree = STRtree(valid_lines)
    intersections = []

    for idx_a, line_a in enumerate(valid_lines):
        original_idx_a = valid_indices[idx_a]
        candidates = tree.query(line_a)

        for idx_b in candidates:
            if idx_b <= idx_a:
                continue

            original_idx_b = valid_indices[idx_b]
            line_b = valid_lines[idx_b]

            if line_a.intersects(line_b):
                intersection = line_a.intersection(line_b)
                if not intersection.is_empty:
                    if hasattr(intersection, 'x'):
                        pts = [[intersection.x, intersection.y]]
                    elif hasattr(intersection, 'geoms'):
                        pts = [[g.x, g.y] for g in intersection.geoms if hasattr(g, 'x')]
                    else:
                        pts = [[intersection.centroid.x, intersection.centroid.y]]

                    for pt in pts:
                        int_point = Point(pt)
                        intersections.append({
                            'path_a': original_idx_a,
                            'path_b': original_idx_b,
                            'point': pt,
                            'is_crossover': is_crossover(int_point, line_a, line_b),
                            'a_has_vertex': point_near_vertex(int_point, line_a, PATH_ANCHOR_TOLERANCE),
                            'b_has_vertex': point_near_vertex(int_point, line_b, PATH_ANCHOR_TOLERANCE)
                        })

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Found {len(intersections)} intersections in {elapsed_ms:.1f}ms")

    return {
        'intersections': intersections,
        'time_ms': elapsed_ms
    }


def build_connection_groups(paths_data: List[Dict], max_dist: float = CLOSE_DIST) -> Dict:
    """
    Build groups of connected paths using Union-Find algorithm.
    Returns which paths belong to which group (for compounding).
    """
    start_time = time.time()

    # First find all connections
    conn_result = find_connections(paths_data, max_dist)
    connections = conn_result['connections']

    # Union-Find for grouping
    n = len(paths_data)
    parent = list(range(n))
    rank = [0] * n

    def find(x):
        if parent[x] != x:
            parent[x] = find(parent[x])
        return parent[x]

    def union(x, y):
        px, py = find(x), find(y)
        if px == py:
            return
        if rank[px] < rank[py]:
            px, py = py, px
        parent[py] = px
        if rank[px] == rank[py]:
            rank[px] += 1

    # Union connected paths
    for conn in connections:
        union(conn['a'], conn['b'])

    # Build groups
    groups = {}
    for i in range(n):
        root = find(i)
        if root not in groups:
            groups[root] = []
        groups[root].append(i)

    # Convert to list format
    group_list = [indices for indices in groups.values() if len(indices) > 1]

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Built {len(group_list)} groups in {elapsed_ms:.1f}ms")

    return {
        'groups': group_list,
        'connections': connections,
        'ignored_anchors': conn_result.get('ignored_anchors', []),
        'time_ms': elapsed_ms
    }


# Main entry point
OPERATIONS = {
    'find_connections': find_connections,
    'detect_intersections': detect_intersections,
    'build_groups': build_connection_groups,
}


def main():
    if len(sys.argv) < 2:
        print("Usage: python geometry_engine.py <operation>", file=sys.stderr)
        print(f"Operations: {', '.join(OPERATIONS.keys())}", file=sys.stderr)
        sys.exit(1)

    operation = sys.argv[1]

    if operation not in OPERATIONS:
        print(f"Unknown operation: {operation}", file=sys.stderr)
        sys.exit(1)

    try:
        # Read input from stdin
        input_data = json.load(sys.stdin)
        paths = input_data.get('paths', [])
        params = input_data.get('params', {})

        log(f"Running {operation} on {len(paths)} paths...")

        # Run the operation
        if params:
            result = OPERATIONS[operation](paths, **params)
        else:
            result = OPERATIONS[operation](paths)

        # Output result as JSON
        json.dump(result, sys.stdout)

    except Exception as e:
        error_result = {'error': str(e), 'operation': operation}
        json.dump(error_result, sys.stdout)
        log(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
