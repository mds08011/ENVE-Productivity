# --- USER: CONFIGURE THESE ---
$DownloadsFolder = "C:\Users\$env:USERNAME\Downloads"
$StagingFolder = "K:\To-Be-Filed" # A central folder to dump files for sorting
# ------------------------------

# 1. Define the file types you care about
$FileTypes = @("*.pdf", "*.dwg", "*.dxf", "*.xlsx", "*.docx")

Write-Host "Checking for files in $DownloadsFolder..."

# 2. Check if the staging folder exists
if (-not (Test-Path $StagingFolder)) {
    Write-Host "Creating staging folder: $StagingFolder"
    New-Item -Path $StagingFolder -ItemType Directory
}

# 3. Find and move the files
Get-ChildItem -Path $DownloadsFolder -Include $FileTypes -Recurse | ForEach-Object {
    $Destination = Join-Path -Path $StagingFolder -ChildPath $_.Name
    
    # Check for file conflicts before moving
    if (Test-Path $Destination) {
        Write-Host "Skipped (name conflict): $($_.Name)" -ForegroundColor Yellow
    } else {
        Move-Item -Path $_.FullName -Destination $StagingFolder
        Write-Host "Moved: $($_.Name)" -ForegroundColor Green
    }
}

Write-Host "Download sorting complete. Files are in $StagingFolder"
Start-Sleep -Seconds 4