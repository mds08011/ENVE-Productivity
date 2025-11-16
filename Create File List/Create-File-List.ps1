# --- IMPROVED FILE LIST GENERATOR ---

# 1. Default to the current folder (where the script is saved)
$sourceFolder = $PSScriptRoot

# Optional: Ask user if they want to scan a different folder
Write-Host "Default scan location: $sourceFolder" -ForegroundColor Cyan
$inputPath = Read-Host "Press ENTER to scan this folder, or paste a new path here"

if ([string]::IsNullOrWhiteSpace($inputPath) -eq $false) {
    # If user typed something, use that instead
    $sourceFolder = $inputPath.Trim('"') # Removes quotes if they copy-pasted path
}

# 2. Set Output File
$outputFile = Join-Path -Path $PSScriptRoot -ChildPath "_Specification_File_List.txt"

# 3. Start
Write-Host "Scanning: $sourceFolder..." -ForegroundColor Yellow
if (Test-Path $outputFile) { Remove-Item $outputFile }

# 4. Get Files and Sort
# We use 'SilentlyContinue' to skip folders we don't have permission for
$allFiles = Get-ChildItem -Path $sourceFolder -Recurse -File -ErrorAction SilentlyContinue | Sort-Object DirectoryName, Name

# 5. Write Output
$currentDirectory = ""
foreach ($file in $allFiles) {
    # Create a clean "Header" for each new subfolder found
    if ($file.DirectoryName -ne $currentDirectory) {
        $currentDirectory = $file.DirectoryName
        
        # Calculate a "Relative Path" so it looks cleaner in the text file
        # e.g., instead of "K:\Projects\...\Specs\Civil", just show "\Civil"
        $relativePath = $currentDirectory.Replace($sourceFolder, "")
        if ($relativePath -eq "") { $relativePath = "ROOT FOLDER" }
        
        Add-Content -Path $outputFile -Value "" 
        Add-Content -Path $outputFile -Value "--- $relativePath ---"
    }
    Add-Content -Path $outputFile -Value "    $($file.Name)"
}

Write-Host "Success! List saved to: $outputFile" -ForegroundColor Green
Start-Sleep -Seconds 4