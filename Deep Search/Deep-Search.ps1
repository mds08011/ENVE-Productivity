# --- DEEP NETWORK SEARCH TOOL ---

# 1. Set the search root (Default to current folder, or change to "K:\")
$SearchRoot = $PSScriptRoot 
# $SearchRoot = "K:\Projects" # Uncomment to make this always search the whole drive

$Keyword = Read-Host "Enter keyword to find (e.g. 'Pump', 'Geotech')"
Write-Host "Searching... (This may take a moment)" -ForegroundColor Yellow

# 2. Find files and output to a GUI Grid View
$results = Get-ChildItem -Path $SearchRoot -Recurse -Filter "*$Keyword*" -ErrorAction SilentlyContinue | 
    Select-Object Name, Directory, LastWriteTime, @{N="Size(KB)";E={[math]::Round($_.Length/1KB,2)}} |
    Out-GridView -Title "Search Results for '$Keyword' - Select and click OK to copy path" -PassThru

# 3. Copy selected path to clipboard
if ($results) {
    $fullPath = Join-Path -Path $results.Directory -ChildPath $results.Name
    Set-Clipboard -Value $fullPath
    Write-Host "Copied path to clipboard: $fullPath" -ForegroundColor Green
}