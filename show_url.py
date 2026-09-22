"""Reads the current cloudflare tunnel URL and shows it."""
from pathlib import Path
import re

log = Path(r"C:\Users\mylap\AppData\Local\Temp\opencode\tunnel_url.txt")
desktop = Path.home() / "Desktop"

url = None
if log.exists():
    for line in log.read_text(encoding="utf-8", errors="ignore").splitlines():
        m = re.search(r'(https://[a-z0-9-]+\.trycloudflare\.com)', line)
        if m:
            url = m.group(1)

if url:
    print(f"Current URL: {url}")
    # Save to desktop for easy access
    (desktop / "EIT_URL.txt").write_text(f"EIT Dashboard URL\n{url}\n\nOpen this link on your mobile.", encoding="utf-8")
    print(f"Saved to Desktop/EIT_URL.txt")
else:
    print("No tunnel URL found. Make sure cloudflared is running.")
