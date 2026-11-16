"""Serves the app on http://localhost:5500 for local testing.

Like `python -m http.server 5500`, but tells the browser to check for a newer copy of every
file on each load (Cache-Control: no-cache), so it never runs an outdated version after an
update. Run from any folder:  python serve.py

Also passes GET requests to Power BI's own servers (*.analysis.windows.net) at
/pbi-proxy?url=…, with the caller's sign-in token. Exports use it to read a report's layout
(colours, drill levels) when the user can't download the report file, as workspace Viewers
can't; browsers refuse to call that Power BI service directly. POST is passed only for the
query a report page runs to draw its charts (/explore/querydata): Viewers' exports use it for
values at a chart's drill level. Nothing is stored.
"""
import http.server
import os
import urllib.error
import urllib.parse
import urllib.request

PORT = 5500
PROXY_PATH = "/pbi-proxy"
PROXY_HOST_SUFFIX = ".analysis.windows.net"
QUERY_PATH = "/explore/querydata"


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == PROXY_PATH:
            self.proxy(urllib.parse.parse_qs(parsed.query).get("url", [""])[0])
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        url = urllib.parse.parse_qs(parsed.query).get("url", [""])[0]
        # Only the query a report page runs to draw its charts
        if parsed.path != PROXY_PATH or not urllib.parse.urlparse(url).path.endswith(QUERY_PATH):
            self.send_error(404)
            return
        length = int(self.headers.get("Content-Length") or 0)
        self.proxy(url, self.rfile.read(length) if length else b"")

    def proxy(self, url, body=None):
        target = urllib.parse.urlparse(url)
        if target.scheme != "https" or not (target.hostname or "").lower().endswith(PROXY_HOST_SUFFIX):
            self.send_error(400, "Only Power BI addresses (*.analysis.windows.net) can be passed through")
            return
        headers = {"Accept": "application/json"}
        for name in ("Authorization", "Content-Type", "ActivityId", "RequestId"):
            if self.headers.get(name):
                headers[name] = self.headers[name]
        try:
            request = urllib.request.Request(url, data=body, headers=headers, method="POST" if body is not None else "GET")
            with urllib.request.urlopen(request, timeout=60) as res:
                status, content_type, body = res.status, res.headers.get("Content-Type", "application/json"), res.read()
        except urllib.error.HTTPError as e:
            status, content_type, body = e.code, e.headers.get("Content-Type", "text/plain"), e.read()
        except (urllib.error.URLError, TimeoutError) as e:
            self.send_error(502, f"Power BI couldn't be reached: {e}")
            return
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print(f"Serving the app on http://localhost:{PORT}  (Ctrl+C to stop)")
    http.server.ThreadingHTTPServer(("localhost", PORT), NoCacheHandler).serve_forever()
