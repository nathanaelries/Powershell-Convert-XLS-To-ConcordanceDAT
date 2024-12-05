# Excel handling utilities
using module '..\config\settings.psm1'

function Quit-Excel {
    param (
        [Parameter(Mandatory=$true)]
        [System.Object]$Excel
    )
    
    try { 
        if ($Excel) {
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
            Remove-Variable -Name Excel -ErrorAction SilentlyContinue
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } catch {}
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
}

function New-ExcelInstance {
    try {
        $excel = New-Object -ComObject Excel.Application
        if (-not $excel) {
            throw "Failed to create Excel instance"
        }
        
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $excel.DisplayStatusBar = $false
        $excel.EnableAnimations = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableLargeOperationAlert = $false
        
        return $excel
    }
    catch {
        throw "Failed to create Excel instance: $($_.Exception.Message)"
    }
}

function Open-ExcelWorkbook {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$true)]
        [System.Object]$Excel
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Excel file not found: $FilePath"
        }
        
        $workbook = $Excel.Workbooks.Open(
            $FilePath,
            0,        # UpdateLinks
            $true,    # ReadOnly
            [Type]::Missing,  # Format
            [Type]::Missing   # Password
        )
        
        if (-not $workbook) {
            throw "Failed to open workbook"
        }
        
        return $workbook
    }
    catch {
        throw "Failed to open workbook: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Quit-Excel, New-ExcelInstance, Open-ExcelWorkbook