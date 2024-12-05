# Main script for Excel to DAT conversion
using module '.\config\settings.ps1'
using module '.\utils\excel-handler.ps1'
using module '.\utils\file-converter.ps1'

# Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFile,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 10000,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxThreads = 4
)

# Update performance settings if provided
$CONFIG.BATCH_SIZE = $BatchSize
$CONFIG.MAX_THREADS = $MaxThreads

# Validate parameters
if (-not (Test-Path $ExcelFile)) {
    throw "Excel file not found: $ExcelFile"
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force
}

# Initialize
$FileNameList = [System.Collections.ArrayList]@()
$excel = New-ExcelInstance
$workbook = Open-ExcelWorkbook -FilePath $ExcelFile -Excel $excel
$excelFileName = $workbook.Name

# Process each worksheet
foreach ($worksheet in $workbook.Worksheets) {
    Write-Progress -Activity "Processing Worksheets" -Status "Converting $($worksheet.Name)"
    
    $tempName = "$excelFileName`_$($worksheet.Name)"
    $tempPath = Join-Path $env:TEMP "$tempName.tmp"
    
    # Save worksheet as temp file
    $worksheet.SaveAs($tempPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText)
    Start-Sleep -Milliseconds 100
    $FileNameList.Add($tempPath)
}

# Close Excel
$excel.Quit()
Quit-Excel $excel

# Process each temp file
foreach ($tempFile in $FileNameList) {
    $fileName = Split-Path $tempFile -Leaf
    $txtFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.txt')
    $datFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.dat')
    $finalDatFile = Join-Path $OutputFolder ($fileName -replace '\.tmp$', '.dat')

    Write-Progress -Activity "Converting Files" -Status "Processing $fileName"

    # Convert temp to txt using optimized streaming
    Convert-ExcelToTemp -WorksheetPath $tempFile -OutputPath $txtFile

    # Convert txt to DAT using parallel processing
    Convert-TempToDAT -InputPath $txtFile -OutputPath $datFile

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
    }
    finally {
        $source.Close()
        $dest.Close()
    }

    # Cleanup temp files
    Remove-Item -Path $tempFile, $txtFile, $datFile -Force
}

Write-Host "Conversion completed successfully. Files saved to: $OutputFolder"