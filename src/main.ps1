# Main script for Excel to DAT conversion
using module '.\config\settings.ps1'
using module '.\utils\excel-handler.ps1'
using module '.\utils\file-converter.ps1'

# Parameters
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
    [string]$ExcelFile,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1000, 1000000)]
    [int]$BatchSize = 10000,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 32)]
    [int]$MaxThreads = 4
)

# Show usage if no parameters provided
if ($MyInvocation.BoundParameters.Count -eq 0) {
    Write-Host "Usage: .\main.ps1 -ExcelFile <path> -OutputFolder <path> [-BatchSize <number>] [-MaxThreads <number>]"
    Write-Host "Example: .\main.ps1 -ExcelFile 'C:\data\file.xlsx' -OutputFolder 'C:\output' -BatchSize 10000 -MaxThreads 4"
    exit
}

# Update performance settings if provided
$CONFIG.BATCH_SIZE = $BatchSize
$CONFIG.MAX_THREADS = $MaxThreads

# Validate output folder
if (-not (Test-Path $OutputFolder)) {
    try {
        New-Item -ItemType Directory -Path $OutputFolder -Force -ErrorAction Stop
        Write-Host "Created output directory: $OutputFolder"
    }
    catch {
        throw "Failed to create output directory: $($_.Exception.Message)"
    }
}

# Initialize
$FileNameList = [System.Collections.ArrayList]@()
try {
    Write-Host "Initializing Excel..."
    $excel = New-ExcelInstance
    $workbook = Open-ExcelWorkbook -FilePath $ExcelFile -Excel $excel
    $excelFileName = $workbook.Name
    Write-Host "Successfully opened Excel workbook: $excelFileName"

    # Process each worksheet
    $worksheetCount = $workbook.Worksheets.Count
    Write-Host "Processing $worksheetCount worksheet(s)..."
    
    foreach ($worksheet in $workbook.Worksheets) {
        Write-Progress -Activity "Processing Worksheets" -Status "Converting $($worksheet.Name)" -PercentComplete (($worksheet.Index / $worksheetCount) * 100)
        
        $tempName = "$excelFileName`_$($worksheet.Name)"
        $tempPath = Join-Path $env:TEMP "$tempName.tmp"
        
        # Save worksheet as temp file
        $worksheet.SaveAs($tempPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText)
        Start-Sleep -Milliseconds 100
        $FileNameList.Add($tempPath)
        Write-Host "  - Processed worksheet: $($worksheet.Name)"
    }
}
catch {
    Write-Error "Error processing Excel file: $($_.Exception.Message)"
    if ($excel) { Quit-Excel $excel }
    exit 1
}
finally {
    # Close Excel
    if ($excel) {
        $excel.Quit()
        Quit-Excel $excel
        Write-Host "Excel instance closed"
    }
}

# Process each temp file
$totalFiles = $FileNameList.Count
$currentFile = 0

foreach ($tempFile in $FileNameList) {
    $currentFile++
    $fileName = Split-Path $tempFile -Leaf
    $txtFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.txt')
    $datFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.dat')
    $finalDatFile = Join-Path $OutputFolder ($fileName -replace '\.tmp$', '.dat')

    Write-Progress -Activity "Converting Files" -Status "Processing $fileName" -PercentComplete (($currentFile / $totalFiles) * 100)
    Write-Host "Converting file $currentFile of $totalFiles: $fileName"

    try {
        # Convert temp to txt using optimized streaming
        Convert-ExcelToTemp -WorksheetPath $tempFile -OutputPath $txtFile
        Write-Host "  - Completed initial conversion"

        # Convert txt to DAT using parallel processing
        Convert-TempToDAT -InputPath $txtFile -OutputPath $datFile
        Write-Host "  - Completed DAT conversion"

        # Copy to final destination using larger buffer
        $buffer = New-Object byte[] $CONFIG.BUFFER_SIZE
        $source = [System.IO.File]::OpenRead($datFile)
        $dest = [System.IO.File]::Create($finalDatFile)
        
        try {
            while ($true) {
                $read = $source.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) { break }
                $dest.Write($buffer, 0, $read)
            }
            Write-Host "  - File saved to output location"
        }
        finally {
            $source.Close()
            $dest.Close()
        }
    }
    catch {
        Write-Error "Error processing file $fileName`: $($_.Exception.Message)"
        continue
    }
    finally {
        # Cleanup temp files
        Remove-Item -Path $tempFile, $txtFile, $datFile -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`nConversion completed successfully!"
Write-Host "Files saved to: $OutputFolder"
