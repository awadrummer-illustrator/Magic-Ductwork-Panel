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


def find_crossovers(paths_data: List[Dict]) -> Dict:
    """
    Find specific segment-segment crossovers between different paths.
    Returns detailed info for splitting segments.
    """
    start_time = time.time()
    
    # 1. Build all segments
    segments = []
    segment_map = [] # (path_idx, seg_idx)
    
    for p_idx, p in enumerate(paths_data):
        points = p.get('points', [])
        if len(points) < 2: continue
        
        # Convert all points to floats once
        pts = [(float(pt['x']), float(pt['y'])) for pt in points]
        
        for i in range(len(pts)-1):
            line = LineString([pts[i], pts[i+1]])
            segments.append(line)
            segment_map.append((p_idx, i))
            
    if not segments:
        return {'crossovers': [], 'time_ms': 0}
        
    # 2. Build STRtree
    tree = STRtree(segments)
    
    crossovers = []
    
    # 3. Query
    for i, seg in enumerate(segments):
        candidates = tree.query(seg)
        path_a_idx, seg_a_idx = segment_map[i]
        
        for j in candidates:
            if j <= i: continue # Avoid duplicates and self-checks
            
            path_b_idx, seg_b_idx = segment_map[j]
            
            # Skip if same path (we only care about crossovers between different paths)
            if path_a_idx == path_b_idx: continue
            
            other_seg = segments[j]
            if seg.intersects(other_seg):
                intersection = seg.intersection(other_seg)
                if not intersection.is_empty:
                    # Handle Point vs MultiPoint etc
                    pts = []
                    if hasattr(intersection, 'x'):
                        pts.append(intersection)
                    elif hasattr(intersection, 'geoms'):
                        pts.extend([g for g in intersection.geoms if hasattr(g, 'x')])
                    else:
                        pts.append(intersection.centroid)
                        
                    for pt in pts:
                        # Check if intersection is at endpoint (vertecies)
                        # We ignore connections at endpoints
                        is_endpoint = (
                            point_near_vertex(pt, seg, 0.1) or 
                            point_near_vertex(pt, other_seg, 0.1)
                        )
                        
                        if not is_endpoint:
                            crossovers.append({
                                'pathIdx': path_a_idx,
                                'segmentIdx': seg_a_idx,
                                'crossingPathIdx': path_b_idx,
                                'crossingSegIdx': seg_b_idx,
                                'point': {'x': pt.x, 'y': pt.y}
                            })

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Found {len(crossovers)} crossovers in {elapsed_ms:.1f}ms")
    
    return {
        'crossovers': crossovers,
        'time_ms': elapsed_ms
    }


def snap_anchors(paths_data: List[Dict], snap_threshold: float = 5.0, locked_points: List[Dict] = None) -> Dict:
    """
    Find anchor points that should snap to nearby segments.
    Uses spatial indexing for O(n log n) instead of O(nÂ²).
    
    locked_points: List of {path_idx, point_idx} dicts for points that should NOT snap.

    Returns snap suggestions: which points should move and where.
    """
    # Build set of locked points for fast lookup
    locked_set = set()
    if locked_points:
        for lp in locked_points:
            locked_set.add((lp.get('path_idx', -1), lp.get('point_idx', -1)))
    start_time = time.time()

    # Build all segments with their path/point indices
    segments = []  # List of (LineString, path_idx, seg_idx)
    segment_lines = []  # Just the LineStrings for spatial index

    for path_idx, path in enumerate(paths_data):
        points = path.get('points', [])
        for seg_idx in range(len(points) - 1):
            p1 = (points[seg_idx]['x'], points[seg_idx]['y'])
            p2 = (points[seg_idx + 1]['x'], points[seg_idx + 1]['y'])
            line = LineString([p1, p2])
            segments.append((line, path_idx, seg_idx))
            segment_lines.append(line)

    if len(segment_lines) == 0:
        return {'snaps': [], 'time_ms': 0}

    # Build spatial index on segments
    tree = STRtree(segment_lines)

    snaps = []
    threshold_sq = snap_threshold * snap_threshold

    for path_idx, path in enumerate(paths_data):
        points = path.get('points', [])
        for pt_idx, pt in enumerate(points):
            # Skip locked points (e.g., ignore marker endpoints)
            if (path_idx, pt_idx) in locked_set:
                continue
                
            point = Point(pt['x'], pt['y'])

            # Query spatial index for nearby segments
            candidates = tree.query(point.buffer(snap_threshold))

            best_snap = None
            best_dist_sq = threshold_sq

            for seg_idx in candidates:
                seg_line, seg_path_idx, _ = segments[seg_idx]

                # Don't snap to own path's segments
                if seg_path_idx == path_idx:
                    continue

                # Find closest point on segment
                closest = seg_line.interpolate(seg_line.project(point))
                dist_sq = (closest.x - pt['x'])**2 + (closest.y - pt['y'])**2

                if dist_sq < best_dist_sq:
                    best_dist_sq = dist_sq
                    best_snap = {
                        'path_idx': path_idx,
                        'point_idx': pt_idx,
                        'snap_to': [closest.x, closest.y],
                        'distance': dist_sq ** 0.5,
                        'target_path_idx': seg_path_idx
                    }

            if best_snap:
                snaps.append(best_snap)

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Found {len(snaps)} snap points in {elapsed_ms:.1f}ms")

    return {
        'snaps': snaps,
        'time_ms': elapsed_ms,
        'path_count': len(paths_data),
        'segment_count': len(segments)
    }


def orthogonalize_paths(paths_data: List[Dict], snap_threshold: float = 5.0,
                        steep_min: float = 17.0, steep_max: float = 70.0,
                        locked_points: List[Dict] = None) -> Dict:
    """
    Full orthogonalization pipeline:
    1. Snap anchors to nearby segments
    2. Orthogonalize segments (make horizontal or vertical)
    3. Return modified path coordinates

    locked_points: List of {path_idx, point_idx} dicts for points that should NOT snap.
    
    This replaces the iterative ExtendScript loop with a single Python call.
    """
    start_time = time.time()
    
    # Build set of locked points for fast lookup
    locked_set = set()
    if locked_points:
        for lp in locked_points:
            locked_set.add((lp.get('path_idx', -1), lp.get('point_idx', -1)))
        log(f"Locked {len(locked_set)} points from snapping (ignore marker endpoints)")

    # Convert to numpy for faster math
    paths = []
    for p in paths_data:
        points = np.array([[pt['x'], pt['y']] for pt in p.get('points', [])])
        paths.append({
            'points': points,
            'layer': p.get('layerName', ''),
            'id': p.get('id', -1)
        })

    changes_made = True
    iteration = 0
    max_iterations = 8
    total_snaps = 0
    total_ortho = 0

    while changes_made and iteration < max_iterations:
        iteration += 1
        changes_made = False

        # Phase 1: Snap anchors
        # Build segments for spatial index
        segment_lines = []
        segment_info = []

        for path_idx, path in enumerate(paths):
            pts = path['points']
            if len(pts) < 2:
                continue
            for seg_idx in range(len(pts) - 1):
                line = LineString([pts[seg_idx], pts[seg_idx + 1]])
                segment_lines.append(line)
                segment_info.append((path_idx, seg_idx))

        if segment_lines:
            tree = STRtree(segment_lines)
            threshold_sq = snap_threshold * snap_threshold

            for path_idx, path in enumerate(paths):
                pts = path['points']
                for pt_idx in range(len(pts)):
                    # Skip locked points (ignore marker endpoints)
                    if (path_idx, pt_idx) in locked_set:
                        continue
                        
                    point = Point(pts[pt_idx])
                    candidates = tree.query(point.buffer(snap_threshold))

                    for seg_idx in candidates:
                        seg_path_idx, _ = segment_info[seg_idx]
                        if seg_path_idx == path_idx:
                            continue

                        seg_line = segment_lines[seg_idx]
                        closest = seg_line.interpolate(seg_line.project(point))
                        dist_sq = (closest.x - pts[pt_idx][0])**2 + (closest.y - pts[pt_idx][1])**2

                        if dist_sq < threshold_sq and dist_sq > 0.01:  # Don't snap if already there
                            pts[pt_idx] = np.array([closest.x, closest.y])
                            changes_made = True
                            total_snaps += 1
                            break

        # Phase 2: Orthogonalize segments
        for path in paths:
            pts = path['points']
            if len(pts) < 2:
                continue

            for i in range(len(pts) - 1):
                p1 = pts[i]
                p2 = pts[i + 1]

                dx = p2[0] - p1[0]
                dy = p2[1] - p1[1]

                if abs(dx) < 0.001 and abs(dy) < 0.001:
                    continue

                # Calculate normalized angle (-90 to 90)
                angle_deg = np.degrees(np.arctan2(dy, dx))
                if angle_deg > 90:
                    angle_deg = angle_deg - 180
                elif angle_deg < -90:
                    angle_deg = angle_deg + 180

                # Check if within 10° of ±45° - snap to exactly 45°
                if (35 <= angle_deg <= 55) or (-55 <= angle_deg <= -35):
                    # Snap to 45° or -45°
                    segment_length = np.sqrt(dx * dx + dy * dy)
                    diagonal_component = segment_length / np.sqrt(2)

                    # Preserve direction signs
                    if angle_deg > 0:
                        # Positive 45° - snap to exactly 45°
                        new_dx = diagonal_component if dx > 0 else -diagonal_component
                        new_dy = diagonal_component if dy > 0 else -diagonal_component
                    else:
                        # Negative 45° - snap to exactly -45°
                        new_dx = diagonal_component if dx > 0 else -diagonal_component
                        new_dy = -diagonal_component if dy < 0 else diagonal_component

                    # Apply the snap
                    pts[i + 1][0] = pts[i][0] + new_dx
                    pts[i + 1][1] = pts[i][1] + new_dy
                    changes_made = True
                    total_ortho += 1
                    continue

                angle = abs(angle_deg)

                # Skip steep angles (non-orthogonal by design)
                if steep_min <= angle <= steep_max:
                    continue

                # Decide whether to make horizontal or vertical
                if abs(dx) > abs(dy):
                    # More horizontal - make perfectly horizontal
                    if pts[i + 1][1] != pts[i][1]:
                        mid_y = (p1[1] + p2[1]) / 2
                        pts[i][1] = mid_y
                        pts[i + 1][1] = mid_y
                        changes_made = True
                        total_ortho += 1
                else:
                    # More vertical - make perfectly vertical
                    if pts[i + 1][0] != pts[i][0]:
                        mid_x = (p1[0] + p2[0]) / 2
                        pts[i][0] = mid_x
                        pts[i + 1][0] = mid_x
                        changes_made = True
                        total_ortho += 1

    # Convert back to output format
    result_paths = []
    for path in paths:
        result_paths.append({
            'id': path['id'],
            'points': [{'x': float(pt[0]), 'y': float(pt[1])} for pt in path['points']]
        })

    elapsed_ms = (time.time() - start_time) * 1000
    log(f"Orthogonalized in {iteration} iterations: {total_snaps} snaps, {total_ortho} ortho changes in {elapsed_ms:.1f}ms")

    return {
        'paths': result_paths,
        'iterations': iteration,
        'total_snaps': total_snaps,
        'total_ortho_changes': total_ortho,
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
    'snap_anchors': snap_anchors,
    'orthogonalize': orthogonalize_paths,
}


def main():
    import argparse

    parser = argparse.ArgumentParser(description='Geometry engine for Magic Ductwork')
    parser.add_argument('operation', help='Operation to perform', choices=list(OPERATIONS.keys()))
    parser.add_argument('--input', '-i', help='Input JSON file (default: stdin)')
    parser.add_argument('--output', '-o', help='Output JSON file (default: stdout)')

    args = parser.parse_args()
    operation = args.operation

    try:
        # Read input from file or stdin
        if args.input:
            with open(args.input, 'r', encoding='utf-8') as f:
                input_data = json.load(f)
        else:
            input_data = json.load(sys.stdin)

        paths = input_data.get('paths', [])
        params = input_data.get('params', {})

        log(f"Running {operation} on {len(paths)} paths...")

        # Run the operation
        if params:
            result = OPERATIONS[operation](paths, **params)
        else:
            result = OPERATIONS[operation](paths)

        # Output result to file or stdout
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(result, f)
        else:
            json.dump(result, sys.stdout)

    except Exception as e:
        error_result = {'error': str(e), 'operation': operation}
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(error_result, f)
        else:
            json.dump(error_result, sys.stdout)
        log(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
