Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python """ & CreateObject("WScript.Shell").CurrentDirectory & "\mobile_app\server.py""", 0, False
WScript.Sleep 4000
WshShell.Run "cmd /c ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=30 -R 80:localhost:8080 nokey@localhost.run > """ & CreateObject("WScript.Shell").CurrentDirectory & "\tunnel_url.txt"" 2>&1", 0, False
