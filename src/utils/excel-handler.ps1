# Excel handling utilities
using module '..\config\settings.ps1'

function Quit-Excel {
    param (
        [Parameter(Mandatory=$true)]
        [System.Object]$Excel
    )
    
    try { 
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {}
    Stop-Process -ErrorAction SilentlyContinue -Name EXCEL -Force 
}

function New-ExcelInstance {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.DisplayStatusBar = $false
    $excel.EnableAnimations = $false
    return $excel
}

function Open-ExcelWorkbook {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [System.Object]$Excel
    )
    
    $Excel.Workbooks.Open($FilePath, 0, $true) # ReadOnly=true for better performance
}

Export-ModuleMember -Function Quit-Excel, New-ExcelInstance, Open-ExcelWorkbook