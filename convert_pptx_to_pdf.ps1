param (
    [string]$pptxFilePath,
    [string]$outputDirectory
)

# Specify the full path to PowerPoint executable
$powerPointPath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"  # Adjust path as necessary

# Check if PowerPoint application exists
if (-not (Test-Path $powerPointPath)) {
    Write-Host "PowerPoint executable not found at: $powerPointPath" -ForegroundColor Red
    exit 1
}

try {
    # Create a new PowerPoint application instance
    $powerPoint = New-Object -ComObject PowerPoint.Application -Strict

    # Open the presentation
    $presentation = $powerPoint.Presentations.Open($pptxFilePath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    # Generate output PDF file path
    $pdfFileName = [System.IO.Path]::GetFileNameWithoutExtension($pptxFilePath) + ".pdf"
    $pdfFilePath = [System.IO.Path]::Combine($outputDirectory, $pdfFileName)

    # Save as PDF
    $presentation.SaveAs($pdfFilePath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)

    # Close the presentation and quit PowerPoint
    $presentation.Close()
    $powerPoint.Quit()

    Write-Host "Conversion of $pptxFilePath to PDF complete. PDF saved at: $pdfFilePath" -ForegroundColor Green
}
catch {
    Write-Host "Error occurred during conversion: $_" -ForegroundColor Red
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
    }
    exit 1
}
