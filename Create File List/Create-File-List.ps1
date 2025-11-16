# --- SCRIPT TO GENERATE A COMBINED LIST OF ALL FILES IN SUBFOLDERS ---

# 1. SET THE FOLDER TO SCAN
$sourceFolder = "K:\SND_WATER\194662101 - Sycuan CWWTP-IPS\Specs\90% Specifications"

# 2. --- MODIFIED: Save output file to the script's own folder ---
$outputFile = $PSScriptRoot + "\_Specification_File_List.txt"

# 3. START SCRIPT
Write-Host "Scanning folder: $sourceFolder"
if (Test-Path $outputFile) { Remove-Item $outputFile }

# 4. GET ALL FILES AND SORT THEM
$allFiles = Get-ChildItem -Path $sourceFolder -Recurse -File | Sort-Object DirectoryName, Name

# 5. FORMAT AND WRITE THE OUTPUT
$currentDirectory = ""
foreach ($file in $allFiles) {
    if ($file.DirectoryName -ne $currentDirectory) {
        $currentDirectory = $file.DirectoryName
        Add-Content -Path $outputFile -Value "" 
        Add-Content -Path $outputFile -Value "--- $($file.Directory.Name) ---"
    }
    Add-Content -Path $outputFile -Value "    $($file.Name)"
}

# 6. FINISH
Write-Host "--- Success! ---"
Write-Host "File list has been created in the script folder: _Specification_File_List.txt"
Start-Sleep -Seconds 10