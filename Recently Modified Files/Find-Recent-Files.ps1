# This script finds all files modified in the last X days in a given folder.

# 1. Get parameters from the user
$FolderPath = Read-Host "Enter the project folder path to scan"
$Days = Read-Host "How many days back to search? (e.g., 2)"

if (-not (Test-Path $FolderPath)) {
    Write-Host "Error: Folder not found. $FolderPath" -ForegroundColor Red
    Start-Sleep -Seconds 3
    exit
}

# 2. Calculate the "cutoff" date
$CutoffDate = (Get-Date).AddDays(-$Days)

Write-Host "Scanning for files modified since $CutoffDate..."

# 3. Get all files, filter by date, sort, and display
Get-ChildItem -Path $FolderPath -Recurse -File | 
    Where-Object { $_.LastWriteTime -gt $CutoffDate } | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -Property LastWriteTime, FullName |
    Format-Table -AutoSize

Read-Host "Press Enter to exit"