import subprocess, time, sys

PORT = 8080
cmd = f"ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=60 -R 80:localhost:{PORT} serveo.net"

while True:
    try:
        proc = subprocess.Popen(cmd.split(), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        for line in proc.stdout:
            print(line.strip(), flush=True)
            if 'serveousercontent.com' in line:
                print(f"\nTUNNEL ACTIVE - Share this URL!", flush=True)
        proc.wait()
    except Exception as e:
        print(f"Error: {e}", flush=True)
    print("Reconnecting in 5s...", flush=True)
    time.sleep(5)
