from http.server import SimpleHTTPRequestHandler, HTTPServer
import sys

class NoCacheHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    httpd = HTTPServer(('0.0.0.0', port), NoCacheHandler)
    print(f"Serving HTTP on 0.0.0.0 port {port} (no-cache enabled)...")
    httpd.serve_forever()
