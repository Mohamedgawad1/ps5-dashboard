$c = Get-Content 'CMT_DASHBOARD\index.html' -Raw

# Search for listOnly
$idx = $c.IndexOf('listOnly')
if ($idx -ge 0) {
    $start = [Math]::Max(0, $idx - 300)
    $len = [Math]::Min(800, $c.Length - $start)
    Write-Output "FOUND listOnly at index $idx"
    Write-Output $c.Substring($start, $len)
} else {
    Write-Output "listOnly NOT FOUND"
}

# Search for list_only
$idx2 = $c.IndexOf('list_only')
if ($idx2 -ge 0) {
    Write-Output "FOUND list_only at index $idx2"
} else {
    Write-Output "list_only NOT FOUND"
}

# Search for List Only (with space)
$idx3 = $c.IndexOf('List Only')
if ($idx3 -ge 0) {
    Write-Output "FOUND List Only at index $idx3"
} else {
    Write-Output "List Only NOT FOUND"
}

# Search for "Fit" or "FIT" in context of display config
$patterns = @('FIT-0001', '"FIT"', "'FIT'", 'fit_type', 'listOnly', 'list_only', 'List Only')
foreach ($p in $patterns) {
    $i = $c.IndexOf($p)
    if ($i -ge 0) {
        $s = [Math]::Max(0, $i - 100)
        $l = [Math]::Min(300, $c.Length - $s)
        Write-Output "--- $p found at $i ---"
        Write-Output $c.Substring($s, $l)
    }
}
