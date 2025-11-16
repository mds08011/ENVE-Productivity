# This script finds and deletes all empty subfolders within its current directory.
# Copy this script into the parent folder you want to clean.

# 1. Get the script's current location
$CurrentFolder = $PSScriptRoot
Write-Host "Scanning for empty folders in: $CurrentFolder"

# 2. Get all subfolders
$allFolders = Get-ChildItem -Path $CurrentFolder -Recurse -Directory

# 3. Sort them by the "depth" (longest path first)
# This is CRITICAL so we delete subfolders *before* their parents
$sortedFolders = $allFolders | Sort-Object @{Expression="($_.FullName -split '\\').Count"} -Descending

$foldersDeleted = 0

# 4. Loop through each folder and check if it's empty
foreach ($folder in $sortedFolders) {
    # Check if the folder has ANY items (files or other folders)
    $itemsInFolder = Get-ChildItem -Path $folder.FullName -ErrorAction SilentlyContinue
    
    if ($null -eq $itemsInFolder) {
        # The folder is empty, so delete it
        Write-Host "Deleting empty folder: $($folder.FullName)" -ForegroundColor Yellow
        Remove-Item -Path $folder.FullName -Force
        $foldersDeleted += 1
    }
}

Write-Host "Complete. Deleted $foldersDeleted empty folders."
Start-Sleep -Seconds 4