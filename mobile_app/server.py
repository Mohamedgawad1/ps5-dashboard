import http.server
import socketserver
import json
import os
import re
import subprocess
import threading
import urllib.parse
import time
from socketserver import ThreadingMixIn

PORT = 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
WIRING_DIR = os.path.join(PARENT_DIR, 'WIRING - MASTER')
TAG_PATTERN = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

refresh_status = {'running': False, 'last': '', 'log': ''}

def scan_single_pdf(filepath):
    try:
        import fitz
        doc = fitz.open(filepath)
        text = ''.join(page.get_text() for page in doc)
        doc.close()
        tags = set(TAG_PATTERN.findall(text))
        return {t.replace(' ', '') for t in tags}
    except Exception:
        return set()

def scan_filename(filename):
    tags = set(TAG_PATTERN.findall(filename))
    return {t.replace(' ', '') for t in tags}

def run_refresh():
    if refresh_status['running']:
        return
    refresh_status['running'] = True
    refresh_status['log'] = 'Starting update...\n'
    try:
        result = subprocess.run(
            ['python', os.path.join(PARENT_DIR, 'daily_update.py')],
            capture_output=True, text=True, timeout=600,
            cwd=PARENT_DIR
        )
        refresh_status['log'] = result.stdout + result.stderr
        refresh_status['last'] = time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        refresh_status['log'] = f'Error: {str(e)}'
    finally:
        refresh_status['running'] = False

class ThreadingHTTPServer(ThreadingMixIn, socketserver.TCPServer):
    daemon_threads = True
    allow_reuse_address = True

class AppHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/data':
            data_path = os.path.join(PARENT_DIR, 'data.json')
            with open(data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
            return

        if path == '/refresh':
            if not refresh_status['running']:
                t = threading.Thread(target=run_refresh, daemon=True)
                t.start()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(json.dumps(refresh_status, ensure_ascii=False).encode('utf-8'))
            return

        if path == '/refresh-status':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(json.dumps(refresh_status, ensure_ascii=False).encode('utf-8'))
            return

        if path.startswith('/pdf/'):
            filename = urllib.parse.unquote(path[5:])
            filepath = os.path.join(WIRING_DIR, filename)
            if not os.path.exists(filepath):
                filepath = os.path.join(PARENT_DIR, filename)
            if not os.path.exists(filepath):
                alt = filename.replace('/', os.sep)
                filepath = os.path.join(WIRING_DIR, alt)
            if not os.path.exists(filepath):
                filepath = os.path.join(PARENT_DIR, alt)
            if os.path.exists(filepath):
                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Disposition', f'inline; filename="{os.path.basename(filepath)}"')
                self.end_headers()
                with open(filepath, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, f'PDF not found: {filename}')
            return

        if path == '/':
            self.path = '/index.html'

        super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/upload':
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'Expected multipart/form-data'}).encode())
                return

            boundary = content_type.split('boundary=')[1].encode()
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            parts = body.split(b'--' + boundary)
            results = []
            uploaded = 0

            for part in parts:
                if b'filename="' not in part:
                    continue
                header_end = part.find(b'\r\n\r\n')
                if header_end == -1:
                    continue
                header = part[:header_end].decode('utf-8', errors='replace')
                file_data = part[header_end+4:]
                if file_data.endswith(b'\r\n'):
                    file_data = file_data[:-2]

                fname_match = re.search(r'filename="([^"]+)"', header)
                if not fname_match:
                    continue
                fname = fname_match.group(1)
                if not fname.lower().endswith('.pdf'):
                    results.append({'name': fname, 'status': 'skipped', 'reason': 'not a PDF'})
                    continue

                save_path = os.path.join(WIRING_DIR, fname)
                with open(save_path, 'wb') as f:
                    f.write(file_data)
                uploaded += 1

                tags_content = scan_filename(fname)
                tags_pdf = scan_single_pdf(save_path)
                all_tags = tags_content | tags_pdf

                mapping_path = os.path.join(PARENT_DIR, 'asset_tag_mapping.json')
                mapping = {}
                if os.path.exists(mapping_path):
                    with open(mapping_path, 'r') as f:
                        mapping = json.load(f)

                new_count = 0
                for t in all_tags:
                    if t not in mapping:
                        new_count += 1
                    mapping[t] = fname

                with open(mapping_path, 'w') as f:
                    json.dump(mapping, f, indent=2)

                results.append({
                    'name': fname,
                    'status': 'ok',
                    'tags': sorted(all_tags),
                    'new_tags': new_count,
                    'total_tags': len(all_tags)
                })

            rebuild_msg = ''
            if uploaded > 0:
                threading.Thread(target=run_refresh, daemon=True).start()
                rebuild_msg = 'Rebuild started in background'

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            resp = {'results': results, 'rebuild': rebuild_msg}
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode('utf-8'))
            return

        self.send_error(404)

    def log_message(self, format, *args):
        pass

print(f"Server running on http://0.0.0.0:{PORT}")
httpd = ThreadingHTTPServer(('0.0.0.0', PORT), AppHandler)
httpd.serve_forever()
