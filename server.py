#!/usr/bin/env python3
"""
Simple HTTP server to serve Excel add-in files and act as a proxy.
Also stores KMAPI settings on the server side.
"""

import http.server
import socketserver
import json
import urllib.request
import urllib.parse
import urllib.error

# Global variable to store settings
SETTINGS = {}

class TestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        if self.path == '/test-api':
            self.handle_test_api()
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/proxy':
            self.handle_proxy()
        elif self.path == '/settings':
            self.handle_settings()
        else:
            self.send_error(404)

    def handle_test_api(self):
        """Test connectivity to constructor.app API"""
        try:
            url = "https://constructor.app/api/platform-kmapi/alive"
            headers = {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }
            
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as response:
                result = response.read().decode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "success",
                    "message": "API connectivity test successful",
                    "response": result
                }).encode())
        except urllib.error.HTTPError as e:
            if e.code == 401:
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "success",
                    "message": "API connectivity test successful (401 expected without API key)",
                    "response": "Unauthorized - this is expected without API key"
                }).encode())
            else:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"API connectivity test failed: {e.code}",
                    "error": str(e)
                }).encode())
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "message": "API connectivity test failed",
                "error": str(e)
            }).encode())

    def handle_proxy(self):
        """Handle proxy requests to external APIs"""
        try:
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            request_data = json.loads(post_data.decode('utf-8'))
            
            url = request_data.get('url')
            method = request_data.get('method', 'POST')
            headers = request_data.get('headers', {})
            body = request_data.get('body', '')
            
            if not url:
                self.send_error(400, "URL is required")
                return
            
            req = urllib.request.Request(url, data=body.encode() if body else None, headers=headers, method=method)
            
            with urllib.request.urlopen(req) as response:
                result = response.read().decode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(result.encode())
                
        except urllib.error.HTTPError as e:
            error_content = e.read().decode('utf-8')
            self.send_response(e.code)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(error_content.encode())
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                "error": str(e)
            }).encode())

    def handle_settings(self):
        """Handle settings storage and retrieval"""
        try:
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            request_data = json.loads(post_data.decode('utf-8'))
            
            action = request_data.get('action')
            
            if action == 'save':
                global SETTINGS
                settings = request_data.get('settings', {})
                SETTINGS.update(settings)
                print(f"Settings saved: {SETTINGS}")
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "success",
                    "message": "Settings saved successfully",
                    "settings": SETTINGS
                }).encode())
                
            elif action == 'get':
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "success",
                    "settings": SETTINGS
                }).encode())
                
            else:
                self.send_error(400, "Invalid action")
                
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "message": str(e)
            }).encode())

def main():
    PORT = 8000
    with socketserver.TCPServer(("", PORT), TestHandler) as httpd:
        print(f"Server running at http://localhost:{PORT}")
        print("Press Ctrl+C to stop")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down...")

if __name__ == "__main__":
    main() 