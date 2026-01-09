"""
Geometry Server for Magic Ductwork Panel

A persistent file-watching server that handles geometry operations.
This eliminates the ~7 second Python startup overhead per call.

Communication happens via files in a watch folder:
- ExtendScript writes: {request_id}_input.json
- Server processes and writes: {request_id}_output.json
- Server maintains: server_status.json (heartbeat)

Usage:
    python geometry_server.py [--watch-folder FOLDER]
"""

import json
import sys
import os
import time
import tempfile
import glob
import threading
import argparse

# Import geometry operations from the engine
from geometry_engine import (
    find_connections,
    detect_intersections,
    build_connection_groups,
    snap_anchors,
    orthogonalize_paths,
    log
)

# Default watch folder - same as ExtendScript uses
DEFAULT_WATCH_FOLDER = os.path.join(tempfile.gettempdir(), "mdux_watch")

# How often to update status file (seconds)
STATUS_UPDATE_INTERVAL = 2.0

# How often to check for new requests (seconds)
POLL_INTERVAL = 0.05  # 50ms


def ensure_watch_folder(folder):
    """Create watch folder if it doesn't exist."""
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder


def write_status_file(folder, running=True):
    """Write server status file for health checks."""
    status_path = os.path.join(folder, "server_status.json")
    status = {
        "status": "running" if running else "stopped",
        "timestamp": int(time.time() * 1000),  # milliseconds
        "pid": os.getpid()
    }
    try:
        with open(status_path, 'w', encoding='utf-8') as f:
            json.dump(status, f)
    except Exception as e:
        log(f"[SERVER] Warning: Could not write status file: {e}")


def status_heartbeat_thread(folder, stop_event):
    """Background thread to update status file periodically."""
    while not stop_event.is_set():
        write_status_file(folder, running=True)
        stop_event.wait(STATUS_UPDATE_INTERVAL)
    # Write final "stopped" status
    write_status_file(folder, running=False)


def process_request(request_data):
    """Process a single geometry request and return result."""
    operation = request_data.get('operation', '')
    request_id = request_data.get('request_id', 'unknown')
    paths = request_data.get('paths', [])
    params = request_data.get('params', {})

    log(f"[SERVER] Processing {operation} (id={request_id}, {len(paths)} paths)")
    start_time = time.time()

    try:
        if operation == 'orthogonalize':
            result = orthogonalize_paths(paths, **params)
        elif operation == 'find_connections':
            result = find_connections(paths, **params)
        elif operation == 'build_groups':
            result = build_connection_groups(paths, **params)
        elif operation == 'detect_intersections':
            result = detect_intersections(paths, **params)
        elif operation == 'snap_anchors':
            result = snap_anchors(paths, **params)
        else:
            return {'error': f'Unknown operation: {operation}'}

        elapsed_ms = (time.time() - start_time) * 1000
        result['time_ms'] = elapsed_ms
        result['request_id'] = request_id
        log(f"[SERVER] {operation} completed in {elapsed_ms:.1f}ms")
        return result

    except Exception as e:
        log(f"[SERVER] Error in {operation}: {e}")
        import traceback
        traceback.print_exc()
        return {'error': str(e), 'request_id': request_id}


def watch_and_process(folder, stop_event):
    """Main loop: watch for input files and process them."""
    log(f"[SERVER] Watching folder: {folder}")
    log(f"[SERVER] Ready to process geometry requests")

    processed_count = 0

    while not stop_event.is_set():
        # Look for input files
        input_pattern = os.path.join(folder, "*_input.json")
        input_files = glob.glob(input_pattern)

        for input_path in input_files:
            try:
                # Extract request ID from filename
                filename = os.path.basename(input_path)
                request_id = filename.replace("_input.json", "")
                output_path = os.path.join(folder, f"{request_id}_output.json")

                # Skip if output already exists (shouldn't happen, but safety check)
                if os.path.exists(output_path):
                    continue

                # Small delay to ensure file is fully written
                time.sleep(0.02)  # 20ms

                # Read input file
                try:
                    with open(input_path, 'r', encoding='utf-8') as f:
                        request_data = json.load(f)
                except (json.JSONDecodeError, ValueError) as e:
                    log(f"[SERVER] JSON error in {filename}: {e}")
                    continue  # Skip, will retry on next poll

                # Process the request
                result = process_request(request_data)

                # Write output file
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f)

                # Clean up input file
                try:
                    os.remove(input_path)
                except:
                    pass

                processed_count += 1

            except Exception as e:
                log(f"[SERVER] Error processing {input_path}: {e}")
                import traceback
                traceback.print_exc()

        # Brief sleep to avoid CPU spin
        time.sleep(POLL_INTERVAL)

    log(f"[SERVER] Shutting down (processed {processed_count} requests)")


def run_server(watch_folder=DEFAULT_WATCH_FOLDER):
    """Run the geometry server."""
    folder = ensure_watch_folder(watch_folder)

    log(f"[SERVER] ========================================")
    log(f"[SERVER] Magic Ductwork Geometry Server")
    log(f"[SERVER] Watch folder: {folder}")
    log(f"[SERVER] ========================================")

    # Event to signal shutdown
    stop_event = threading.Event()

    # Start status heartbeat thread
    heartbeat = threading.Thread(
        target=status_heartbeat_thread,
        args=(folder, stop_event),
        daemon=True
    )
    heartbeat.start()

    # Write initial status
    write_status_file(folder, running=True)

    print(f"READY:{folder}", flush=True)  # Signal to parent process

    try:
        watch_and_process(folder, stop_event)
    except KeyboardInterrupt:
        log("[SERVER] Received shutdown signal")
    finally:
        stop_event.set()
        # Give heartbeat thread time to write final status
        time.sleep(0.5)
        log("[SERVER] Server stopped")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Geometry server for Magic Ductwork')
    parser.add_argument('--watch-folder', '-w', type=str, default=DEFAULT_WATCH_FOLDER,
                        help=f'Folder to watch for requests (default: {DEFAULT_WATCH_FOLDER})')
    parser.add_argument('--port', '-p', type=int, help='Ignored (for backwards compatibility)')
    args = parser.parse_args()

    run_server(args.watch_folder)
