import http.server
import socketserver
import os
import json
import urllib.parse
from sync_excel import append_tender_to_excel, export_to_json, EXCEL_FILE

PORT = 8000

class EnhancedCORSRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        # Redirect root to main Proposals dashboard HTML
        if self.path in ['/', '/index.html', '/dashboard.html']:
            target = 'ADM_APP_ATI_000%20Proposals.html' if os.path.exists('ADM_APP_ATI_000 Proposals.html') else 'Offer%20Dashboard.html'
            self.send_response(302)
            self.send_header('Location', '/' + target)
            self.send_header('Cache-Control', 'no-store, must-revalidate')
            self.end_headers()
            return
        super().do_GET()

    def do_POST(self):
        if self.path == '/api/save_tender':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            try:
                data = json.loads(body.decode('utf-8'))
                new_row = append_tender_to_excel(data)
                export_to_json()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'ok', 'row': new_row}).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    print(f"================================================================")
    print(f"        ATI Offer & Tender Tracker Development Server           ")
    print(f"================================================================")
    print(f"Serving Directory: {os.getcwd()}")
    print(f"Dashboard URL:     http://localhost:{PORT}/")
    print(f"Excel File URL:    http://localhost:{PORT}/{urllib.parse.quote(EXCEL_FILE)}")
    print(f"Direct Browser:    You can also open 'Offer Dashboard.html' directly")
    print(f"                   in Chrome/Edge with NO Python server needed!")
    print(f"================================================================")
    print("Press CTRL+C in this window to stop the server.\n")

    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), EnhancedCORSRequestHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped safely.")
