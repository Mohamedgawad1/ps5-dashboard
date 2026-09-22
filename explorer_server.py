"""
PS5 Explorer Sync Server
Serves subsystem_explorer.html + saves platform edits (notes / row colors)
to _explorer_state.json -> picked up by punch_itr_explorer.py & dpr_dashboard.py
on every daily Excel rebuild.

Usage: python explorer_server.py   (then open http://localhost:8090)
"""
import json
import os
import threading
import time
import webbrowser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

PORT = 8090
DL_SUB = os.path.join(os.path.expanduser('~'),
                      'Downloads', 'PS5 - CPP AGI Completion Progress Dashboard_files')
STATE_FILE = os.path.join(DL_SUB, '_explorer_state.json')

_lock = threading.Lock()


def load_state():
    try:
        with open(STATE_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {'notes': {}, 'colors': {}}


def save_state(st):
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(st, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE_FILE)




class Handler(SimpleHTTPRequestHandler):
    timeout = 45
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=DL_SUB, **kw)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def log_message(self, fmt, *args):
        pass

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)


    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        if self.path.startswith('/api/state'):
            with _lock:
                return self._send_json(load_state())
            return None
        return super().do_GET()

    def do_POST(self):
        if not self.path.startswith('/api/state'):
            return self._send_json({'ok': False, 'err': 'not found'}, 404)
        try:
            n = int(self.headers.get('Content-Length', 0))
            payload = json.loads(self.rfile.read(n).decode('utf-8'))
        except Exception as e:
            return self._send_json({'ok': False, 'err': str(e)}, 400)

        sheet = str(payload.get('sheet', '')).strip()
        rid = str(payload.get('id', '')).strip()
        if sheet not in ('PUNCH LIST', 'ITR LIST', 'RFC PROGRESS') or not rid:
            return self._send_json({'ok': False, 'err': 'bad sheet/id'}, 400)

        with _lock:
            st = load_state()
            col = str(payload.get('col', '')).strip()
            if col:
                val = payload.get('value')
                val = '' if val is None else str(val).strip()
                cells = st.setdefault('cells', {}).setdefault(sheet, {})
                cc = cells.setdefault(rid, {})
                if val:
                    cc[col] = val
                else:
                    cc.pop(col, None)
                if not cc:
                    cells.pop(rid, None)
                save_state(st)
                return self._send_json({'ok': True})
            fld = str(payload.get('field', '')).strip()
            if fld:
                if fld not in ('eacop', 'cpeit', 'cpp1'):
                    return self._send_json({'ok': False, 'err': 'bad field'}, 400)
                val = str(payload.get('note', '')).strip()
                cell = st.setdefault('rmk', {}).setdefault(rid, {})
                if val:
                    cell[fld] = val
                else:
                    cell.pop(fld, None)
                save_state(st)
                return self._send_json({'ok': True})
            if 'note' in payload:
                notes = st.setdefault('notes', {}).setdefault(sheet, {})
                val = str(payload['note']).strip()
                if val:
                    notes[rid] = val
                else:
                    notes.pop(rid, None)
            if 'color' in payload:
                colors = st.setdefault('colors', {}).setdefault(sheet, {})
                col = str(payload['color']).strip()
                if col:
                    colors[rid] = col
                else:
                    colors.pop(rid, None)
            save_state(st)
        return self._send_json({'ok': True})


def open_browser():
    time.sleep(0.8)
    webbrowser.open(f'http://localhost:{PORT}/subsystem_explorer.html')


if __name__ == '__main__':
    import socket

    os.makedirs(DL_SUB, exist_ok=True)

    def port_free(p):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('', p))
                return True
            except OSError:
                return False

    url = f'http://localhost:{PORT}/subsystem_explorer.html'
    if not port_free(PORT):
        print(f'Port {PORT} busy - platform already running, opening browser...')
        webbrowser.open(url)
        raise SystemExit(0)

    print(f'PS5 Explorer Sync Server -> {url}')
    print(f'State file: {STATE_FILE}')
    print('Press Ctrl+C to stop')
    threading.Thread(target=open_browser, daemon=True).start()
    ThreadingHTTPServer(('', PORT), Handler).serve_forever()
