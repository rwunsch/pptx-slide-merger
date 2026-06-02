"""
Reviewer save-server for the pptx-slide-merger review viewer.

The default `python3 -m http.server` is static — it serves the generated viewer
but rejects writes, so the reviewer can only Export the comments and re-drop the
file by hand. This module mirrors the STACK deck's `review-server.mjs`: it serves
the generated viewer directory AND accepts the reviewer's auto-saves.

Reviewer Mode does  PUT (or POST) /review-comments.json  with the comments JSON;
this server validates it and writes the file straight into the viewer folder, so
auto-save persists to disk with no manual Export step.

Run it standalone against a generated viewer dir:

    python3 -m pptx_slide_merger.review_server <viewer-dir> [--port 8000]

or use the `serve-review.py` shim that `build_review_viewer` drops into the
viewer dir, or the `--serve` flag on `pptx-merge review`.
"""

from __future__ import annotations

import argparse
import json
import sys
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

COMMENTS_FILE = "review-comments.json"
_MAX_BODY = 8 * 1024 * 1024  # 8 MB guard, matches review-server.mjs


class _ReviewRequestHandler(SimpleHTTPRequestHandler):
    """Static file server + PUT/POST review-comments.json -> write to disk."""

    # `directory=` is bound via functools.partial in serve(); SimpleHTTPRequestHandler
    # honours it for GET/HEAD. We reuse self.directory for the write target too.

    def _handle_comment_write(self):
        path = (self.path.split("?", 1)[0])
        if not path.endswith("/" + COMMENTS_FILE) and not path.endswith(COMMENTS_FILE):
            self.send_error(404, "not found")
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
        except (TypeError, ValueError):
            length = 0
        if length <= 0 or length > _MAX_BODY:
            self._json(400, {"ok": False, "error": "bad length"})
            return
        body = self.rfile.read(length)
        try:
            json.loads(body)  # validate
        except (ValueError, UnicodeDecodeError):
            self._json(400, {"ok": False, "error": "invalid json"})
            return
        target = Path(self.directory) / COMMENTS_FILE
        try:
            target.write_bytes(body)
        except OSError as exc:
            self._json(500, {"ok": False, "error": str(exc)})
            return
        sys.stderr.write(f"[review] saved {COMMENTS_FILE} ({len(body)} bytes)\n")
        self._json(200, {"ok": True})

    def do_PUT(self):  # noqa: N802 (http.server casing)
        self._handle_comment_write()

    def do_POST(self):  # noqa: N802
        self._handle_comment_write()

    def do_OPTIONS(self):  # noqa: N802 (CORS preflight, just in case)
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,PUT,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def end_headers(self):
        # No-cache so review-comments.json / slides.json reloads are always fresh.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _json(self, status, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):  # quieter than the default
        sys.stderr.write("[review] " + (fmt % args) + "\n")


def serve(directory: Path, port: int = 8000, host: str = "127.0.0.1",
          verbose: bool = True) -> None:
    """Serve `directory` and accept PUT/POST review-comments.json writes."""
    directory = Path(directory).resolve()
    if not directory.is_dir():
        raise NotADirectoryError(f"{directory} is not a directory")
    handler = partial(_ReviewRequestHandler, directory=str(directory))
    httpd = ThreadingHTTPServer((host, port), handler)
    if verbose:
        actual = httpd.server_address[1]
        print(f"Review save-server on http://{host}:{actual}/  "
              f"(PUT {COMMENTS_FILE} persists into {directory})")
        print("Open the URL, toggle Review (or press R), click slides to comment. "
              "Ctrl+C to stop.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        if verbose:
            print("\nReview save-server stopped.")
    finally:
        httpd.server_close()


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Serve a pptx review viewer dir and persist reviewer auto-saves.")
    ap.add_argument("directory", type=Path, nargs="?", default=Path("."),
                    help="Viewer directory (the folder with index.html). Default: cwd")
    ap.add_argument("--port", type=int, default=8000, help="Port (default: 8000)")
    ap.add_argument("--host", default="127.0.0.1", help="Host (default: 127.0.0.1)")
    args = ap.parse_args(argv)
    serve(args.directory, port=args.port, host=args.host)


if __name__ == "__main__":
    main()
