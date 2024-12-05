# Main script for Excel to DAT conversion
[CmdletBinding()]
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

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

# Import required modules
$modulePath = $PSScriptRoot
Import-Module (Join-Path $modulePath "config\settings.psm1") -Force
Import-Module (Join-Path $modulePath "utils\excel-handler.psm1") -Force
Import-Module (Join-Path $modulePath "utils\file-converter.psm1") -Force
Import-Module (Join-Path $modulePath "utils\progress-handler.psm1") -Force
Import-Module (Join-Path $modulePath "utils\cleanup-handler.psm1") -Force

# Update performance settings
$CONFIG.BATCH_SIZE = $BatchSize
$CONFIG.MAX_THREADS = $MaxThreads

# Validate and create output folder if needed
if (-not (Test-Path $OutputFolder)) {
    try {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-Host "Created output directory: $OutputFolder"
    }
    catch {
        throw "Failed to create output directory: $($_.Exception.Message)"
    }
}

# Initialize
$FileNameList = [System.Collections.ArrayList]::new()
$excel = $null

try {
    Write-Host "Initializing Excel..."
    $excel = New-ExcelInstance
    
    $workbook = Open-ExcelWorkbook -FilePath $ExcelFile -Excel $excel
    $excelFileName = $workbook.Name
    Write-Host "Successfully opened Excel workbook: $excelFileName"

    # Process worksheets
    $worksheetCount = $workbook.Worksheets.Count
    Write-Host "Processing $worksheetCount worksheet(s)..."
    
    foreach ($worksheet in $workbook.Worksheets) {
        $worksheetName = $worksheet.Name
        Write-ProcessProgress -Activity "Processing Worksheets" `
                            -Status "Converting $worksheetName" `
                            -Current $worksheet.Index `
                            -Total $worksheetCount
        
        $tempName = "{0}_{1}" -f $excelFileName, $worksheetName
        $tempPath = Join-Path $env:TEMP "$tempName.tmp"
        
        $worksheet.SaveAs($tempPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText)
        Start-Sleep -Milliseconds 100
        [void]$FileNameList.Add($tempPath)
        Write-Host "  - Processed worksheet: $worksheetName"
    }
}
catch {
    Write-Error "Error processing Excel file: $($_.Exception.Message)"
    if ($excel) { Quit-Excel $excel }
    exit 1
}
finally {
    if ($excel) {
        $excel.Quit()
        Quit-Excel $excel
        Write-Host "Excel instance closed"
    }
}

# Process temp files
$currentFile = 0
$totalFiles = $FileNameList.Count

foreach ($tempFile in $FileNameList) {
    $currentFile++
    $fileName = Split-Path $tempFile -Leaf
    $txtFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.txt')
    $datFile = Join-Path $env:TEMP ($fileName -replace '\.tmp$', '.dat')
    $finalDatFile = Join-Path $OutputFolder ($fileName -replace '\.tmp$', '.dat')

    Write-ProcessProgress -Activity "Converting Files" `
                        -Status "Processing $fileName" `
                        -Current $currentFile `
                        -Total $totalFiles
                        
    Write-Host ("Converting file {0} of {1}: {2}" -f $currentFile, $totalFiles, $fileName)

    try {
        Convert-ExcelToTemp -WorksheetPath $tempFile -OutputPath $txtFile
        Write-Host "  - Completed initial conversion"

        Convert-TempToDAT -InputPath $txtFile -OutputPath $datFile
        Write-Host "  - Completed DAT conversion"

        Copy-Item -Path $datFile -Destination $finalDatFile -Force
        Write-Host "  - File saved to output location"
    }
    catch {
        Write-Error ("Error processing file {0}: {1}" -f $fileName, $_.Exception.Message)
        continue
    }
    finally {
        Remove-TempFiles -FilePaths @($tempFile, $txtFile, $datFile)
    }
}

Write-Host "`nConversion completed successfully!"
Write-Host "Files saved to: $OutputFolder"