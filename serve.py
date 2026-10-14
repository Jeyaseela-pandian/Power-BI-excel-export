"""Serves the app on http://localhost:5500 for local testing.

Like `python -m http.server 5500`, but tells the browser to check for a newer copy of every
file on each load (Cache-Control: no-cache), so it never runs an outdated version after an
update. Run from any folder:  python serve.py
"""
import http.server
import os

PORT = 5500


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print(f"Serving the app on http://localhost:{PORT}  (Ctrl+C to stop)")
    http.server.ThreadingHTTPServer(("localhost", PORT), NoCacheHandler).serve_forever()
