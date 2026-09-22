$c = Get-Content 'CMT_DASHBOARD\index.html' -Raw
$idx = $c.IndexOf('FIT-0001-CJ01')
if ($idx -ge 0) {
    $start = [Math]::Max(0, $idx - 300)
    $len = [Math]::Min(800, $c.Length - $start)
    Write-Output "FOUND at index $idx"
    Write-Output $c.Substring($start, $len)
} else {
    Write-Output "NOT FOUND in CMT_DASHBOARD\index.html"
}
