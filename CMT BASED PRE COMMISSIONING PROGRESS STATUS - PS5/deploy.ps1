$token = (Get-Content "C:\Users\mylap\OneDrive\Desktop\dashboard\github_token.txt" -Raw).Trim()
$repo = "mohamedgawad1/CMT-COMMISSIONING-PROGRESS-STATUS-PS5"
$path = "index.html"
$file = "$PSScriptRoot\index.html"

$content = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file))

try {
    $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$path" -Headers @{Authorization="token $token"} -ErrorAction Stop
    $sha = $existing.sha
} catch { $sha = $null }

$body = @{
    message = "auto update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    content = $content
    branch = "main"
}
if ($sha) { $body.sha = $sha }

$json = $body | ConvertTo-Json
Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$path" -Method Put -Headers @{Authorization="token $token"} -Body $json -ContentType "application/json" | Out-Null
Write-Host "index.html uploaded to $repo"
