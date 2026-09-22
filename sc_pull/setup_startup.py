import os, sys, subprocess

ws = os.getcwd()
startup = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows",
                       "Start Menu", "Programs", "Startup")
vbs = os.path.join(startup, "PS5-ITR-Live.vbs")

lines = [
    'Set sh = CreateObject("Wscript.Shell")',
    'sh.CurrentDirectory = "' + ws + '"',
    'sh.Run "python -u -W ignore sc_pull\\itr_online_sync.py --loop", 0, False',
]
open(vbs, "w", encoding="utf-8").write("\r\n".join(lines) + "\r\n")
print("VBS written:", vbs)
print("startup exists:", os.path.isdir(startup))

# launch it now too so it takes effect immediately
subprocess.Popen(["wscript.exe", vbs])
print("launched via wscript")
