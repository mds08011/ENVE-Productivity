# --- USER: CONFIGURE THIS ---
$ProjectRoot = "K:\Projects"  # The main folder containing all your projects
# ------------------------------

Write-Host "Calculating sizes of all folders in $ProjectRoot... This may take time."

# 1. Get all project folders (but not files) in the root
$projectFolders = Get-ChildItem -Path $ProjectRoot -Directory

$results = @()

# 2. Loop through each project
foreach ($folder in $projectFolders) {
    Write-Host "Scanning: $($folder.Name)"
    
    # 3. Get all files in this project, recursively
    $files = Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue
    
    # 4. Measure the total size
    $size = $files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
    
    # 5. Store the results
    $results += [PSCustomObject]@{
        ProjectName = $folder.Name
        SizeInGB = [math]::Round($size.Sum / 1GB, 2)
        FileCount = $size.Count
    }
}

# 6. Display the final report, sorted by size (largest first)
Write-Host "--- Project Size Report ---" -ForegroundColor Cyan
$results | Sort-Object -Property SizeInGB -Descending | Format-Table -AutoSize

Read-Host "Press Enter to exit"