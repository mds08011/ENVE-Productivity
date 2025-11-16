# --- Script to Convert Word Docs to PDF ---

# Get the directory where the script is located
$scriptPath = Get-Location

# Define the PDF format code for Word
$wdFormatPDF = 17

# Create an instance of the Word Application (it will run hidden)
Write-Host "Starting Word application..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Find all .doc and .docx files in the current directory
$docFiles = Get-ChildItem -Path $scriptPath -Include "*.doc", "*.docx" -Recurse

# Check if any documents were found
if ($docFiles.Count -eq 0) {
    Write-Host "No .doc or .docx files found in this directory."
} else {
    Write-Host "Found $($docFiles.Count) documents to convert."

    # Loop through each document
    foreach ($docFile in $docFiles) {
        $docPath = $docFile.FullName
        $pdfPath = $docPath.Replace($docFile.Extension, ".pdf")

        # Check if a PDF with the same name already exists
        if (Test-Path $pdfPath) {
            Write-Host "Skipping '$($docFile.Name)' because a PDF version already exists."
            continue
        }

        # Open the Word document
        Write-Host "Converting '$($docFile.Name)'..."
        $document = $word.Documents.Open($docPath)

        # Save the document as a PDF
        $document.SaveAs($pdfPath, $wdFormatPDF)

        # Close the document without saving changes
        $document.Close($false)
    }
}

# Quit the Word application and clean up
Write-Host "Closing Word application."
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "Conversion complete!"