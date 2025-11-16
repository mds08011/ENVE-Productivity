# This script deletes common auxiliary files from hydraulic models.
# Copy this script into the model folder you want to clean.

# --- USER: CONFIGURE THIS ---
# Add or remove file extensions you want to delete.
$ExtensionsToDelete = @(
    "*.bak",
    "*.out",
    "*.rpt",
    "*.log",
    "*.tmp",
    "*.err",
    "*.dwh",
    "*.sqlite-journal"
)
# ------------------------------

# 1. Get the script's current location
$CurrentFolder = $PSScriptRoot
Write-Host "Scanning for files to clean in: $CurrentFolder" -ForegroundColor Yellow

# 2. Find all files matching the extensions
$filesToDelete = Get-ChildItem -Path $CurrentFolder -Include $ExtensionsToDelete -Recurse

if ($filesToDelete.Count -eq 0) {
    Write-Host "No temporary files found. Folder is already clean." -ForegroundColor Green
    Start-Sleep -Seconds 3
    exit
}

# 3. List the files and ask for confirmation
Write-Host "The following $($filesToDelete.Count) files will be DELETED:"
$filesToDelete | ForEach-Object { Write-Host " - $($_.Name)" }

$Confirmation = Read-Host "Are you sure you want to delete these files? (y/n)"

# 4. Delete the files
if ($Confirmation -eq 'y') {
    $filesToDelete | Remove-Item -Force
    Write-Host "Cleanup complete." -ForegroundColor Green
} else {
    Write-Host "Operation cancelled."
}
Start-Sleep -Seconds 4