$ErrorActionPreference='SilentlyContinue'
$root="C:\Users\mylap\OneDrive\Desktop\dashboard"
$log="$env:TEMP\opencode\share_tunnel.log"
New-Item -ItemType Directory -Force -Path "$env:TEMP\opencode" | Out-Null
$t=Test-NetConnection -ComputerName localhost -Port 8090 -WarningAction SilentlyContinue
if(-not $t.TcpTestSucceeded){
  Start-Process -FilePath python -ArgumentList "`"$root\explorer_server.py`"" -WindowStyle Minimized
  Start-Sleep -Seconds 3
}
if(Get-Process cloudflared -ErrorAction SilentlyContinue){ Stop-Process -Name cloudflared -Force }
Start-Sleep -Seconds 1
$proc=Start-Process -FilePath "$root\cloudflared.exe" -ArgumentList 'tunnel','--url','http://localhost:8090','--no-autoupdate' -RedirectStandardOutput $log -RedirectStandardError "$env:TEMP\opencode\share_err.log" -PassThru -WindowStyle Hidden
$url=$null
for($i=0;$i -lt 40 -and -not $url;$i++){
  Start-Sleep -Milliseconds 700
  $txt=(Get-Content $log -Raw -ErrorAction SilentlyContinue)+(Get-Content "$env:TEMP\opencode\share_err.log" -Raw -ErrorAction SilentlyContinue)
  if($txt -match 'https://[a-z0-9\-]+\.trycloudflare\.com'){$url=$Matches[0]}
}
if($url){
  Set-Content -Path "$env:USERPROFILE\Desktop\PS5 SHARE LINK.txt" -Value "PS5 PLATFORM - SHARE THIS LINK:`r`n$url`r`n(works while this PC is on)`r`n" -Encoding UTF8
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show("SHARE THIS LINK WITH THE ENGINEERS:`r`n`r`n$url`r`n`r`nA copy was saved to Desktop (PS5 SHARE LINK.txt)`r`nKeep this PC ON while they work.","PS5 PLATFORM - LIVE SHARE",0,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}else{
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show("Tunnel did not start. Check internet connection and run again.","PS5 SHARE",0,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}
