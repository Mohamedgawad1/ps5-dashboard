import json, os, time, datetime, subprocess, threading
from http.server import HTTPServer, BaseHTTPRequestHandler

RECIPIENTS = [
    "+255684064292",
    "+255677114429",
    "+255674643758",
    "+255699065843",
]

BASE = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
ADB = r"C:\adb\platform-tools\adb"
PHONE = "swjrci9xythaizuk"

def send_sms(number, msg):
    body_esc = msg.replace('\n', '\\n')
    subprocess.run([ADB, '-s', PHONE, 'shell', 'am', 'start',
        '-a', 'android.intent.action.VIEW',
        '-d', f'sms:{number}',
        '--es', 'sms_body', body_esc])
    print(f"  Opened for {number}")

def push_data():
    sms_path = os.path.join(BASE, 'sms_data.json')
    subprocess.run([ADB, '-s', PHONE, 'push', sms_path, '/sdcard/sms_data.json'],
        capture_output=True)
    print("  Data pushed to phone")

def build_msg():
    sms_path = os.path.join(BASE, 'sms_data.json')
    try:
        with open(sms_path, 'r') as f:
            data = json.load(f)
    except:
        data = {}
    tc = data.get('today_closed', 0)
    ts = data.get('today_submitted', 0)
    e = data.get('e', {})
    i = data.get('i', {})
    t = data.get('t', {})
    today = datetime.datetime.now().strftime('%Y-%m-%d')
    if data.get('date', '') != today:
        tc = 0; ts = 0
    return (
        f"ITR Today Close {tc} Submit {ts}\n"
        f"E Total {e.get('total',0)} Closed {e.get('closed',0)} Open {e.get('open',0)}\n"
        f"I Total {i.get('total',0)} Closed {i.get('closed',0)} Open {i.get('open',0)}\n"
        f"T Total {t.get('total',0)} Closed {t.get('closed',0)} Open {t.get('open',0)}"
    )

def send_status():
    now = datetime.datetime.now()
    push_data()
    msg = build_msg()
    print(f"\n[{now.strftime('%H:%M')}] Opening Message for all recipients...")
    print(msg)
    print("  -> Tap SEND on each one!")
    for r in RECIPIENTS:
        send_sms(r, msg)
        time.sleep(5)

def send_to_number(number, msg):
    body_esc = msg.replace('\n', '\\n')
    subprocess.run([ADB, '-s', PHONE, 'shell', 'am', 'start',
        '-a', 'android.intent.action.VIEW',
        '-d', f'sms:{number}',
        '--es', 'sms_body', body_esc])
    print(f"  Opened for {number}")

class SMSHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/send-sms':
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.end_headers()
            self.wfile.write(b'ok')
            print(f"\n[HTTP] Send triggered from dashboard")
            threading.Thread(target=send_status, daemon=True).start()
        elif self.path.startswith('/send-to'):
            from urllib.parse import urlparse, parse_qs
            qs = parse_qs(urlparse(self.path).query)
            number = qs.get('number', [None])[0]
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.end_headers()
            if not number:
                self.wfile.write(b'missing number')
                print(f"\n[HTTP] send-to missing number")
            else:
                self.wfile.write(b'ok')
                print(f"\n[HTTP] send-to {number} triggered")
                threading.Thread(target=send_to_number, args=(number, build_msg()), daemon=True).start()
        else:
            self.send_response(404)
            self.end_headers()
    def log_message(self, *a): pass

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == 'send':
        if len(sys.argv) < 3 or sys.argv[2] != 'Adam@goda1':
            print("ERROR: Password required. Usage: python sms_laptop.py send <password>")
            sys.exit(1)
        push_data()
        send_status()
        sys.exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == 'serve':
        push_data()
        httpd = HTTPServer(('127.0.0.1', 8765), SMSHandler)
        t = threading.Thread(target=httpd.serve_forever, daemon=True)
        t.start()
        print(f"SMS server running on http://127.0.0.1:8765")
        print(f"  Dashboard SMS button will call this server (opens compose on phone, tap SEND).")
        print(f"  (Or run: python sms_laptop.py send <password>  to send now)")
        while True:
            time.sleep(30)
