# This script renames JPG/JPEG files based on their 'Date Taken' EXIF metadata.
# Copy this script into the folder of photos you want to rename.

# 1. Ask user for a prefix
$Prefix = Read-Host "Enter a prefix for all photos (e.g., 'Pump-Station-Slab')"
$Prefix = $Prefix.Trim() + "_"

# 2. Load the required .NET assembly
Add-Type -Assembly System.Drawing

# 3. Get all JPG files
$imageFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.jpg", "*.jpeg"

Write-Host "Processing $($imageFiles.Count) photos..."

# 4. Loop through each image
foreach ($file in $imageFiles) {
    try {
        # 5. Open the image file and read its metadata
        $image = New-Object System.Drawing.Bitmap($file.FullName)
        
        # 36867 is the EXIF ID for "Date Taken"
        $dateTakenProperty = $image.GetPropertyItem(36867) 
        
        # Convert the property to a string
        $encoding = New-Object System.Text.ASCIIEncoding
        $dateTakenStr = $encoding.GetString($dateTakenProperty.Value, 0, $dateTakenProperty.Len - 1)
        
        # 6. Format the date string into something file-safe
        # From '2025:11:02 14:30:15' to '2025-11-02_14-30-15'
        $formattedDate = $dateTakenStr -replace ':', '-'
        $formattedDate = $formattedDate -replace ' ', '_'
        
        # 7. Create the new file name
        $NewName = $formattedDate + "_" + $Prefix + $file.Name
        
        # 8. IMPORTANT: Release the file lock before renaming
        $image.Dispose()
        
        # 9. Rename the file
        Rename-Item -Path $file.FullName -NewName $NewName
        Write-Host "Renamed $($file.Name) to $NewName" -ForegroundColor Green
    }
    catch {
        # 10. Handle errors (e.g., photo has no EXIF data)
        Write-Host "Could not process $($file.Name). No 'Date Taken' data?" -ForegroundColor Red
        if ($image) { $image.Dispose() } # Ensure file lock is released on error
    }
}

Write-Host "Photo renaming complete."
Start-Sleep -Seconds 4