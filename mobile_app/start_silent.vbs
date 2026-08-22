Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python ""C:\Users\mylap\OneDrive\Desktop\dashboard\mobile_app\server.py""", 0, False
WshShell.Run "python ""C:\Users\mylap\OneDrive\Desktop\dashboard\mobile_app\tunnel.py""", 0, False
