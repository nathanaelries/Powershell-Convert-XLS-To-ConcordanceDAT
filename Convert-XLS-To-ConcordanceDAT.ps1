# User Inputs
$excelFile = "\\path\to\excel.xls"
$DatLoc = "\\Path\to\output\folder\"
# END User Inputs

$pipe = [char](20)
$thorne = [char](254)
$tempfile = [guid]::NewGuid()
$filename = $null
$CSVData = $null
$script:tab = [char](9)

# close the Excel without saving, three ways to close in case Excel resists being killed
function Quit-Excel ($Excel) {
try{$Excel.Quit()}catch{}
try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)}catch{}
Stop-Process -ErrorAction SilentlyContinue -Name EXCEL -Force 
}
Quit-Excel

[System.Collections.ArrayList]$FileNameList = @()
function EXportWSToCSV ($excelFile){
    foreach ($ws in $WB.Worksheets){
        $n = $excelFileName +"_"+ $ws.Name
        $WS.SaveAs((("$env:TEMP\" + $n + ".tmp")-replace '"',''), [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText)
        Start-Sleep -Milliseconds 100
        $Script:FileNameList += (("$env:TEMP\" + $n + ".tmp")-replace '"','')
    }
    $E.Quit()
    Quit-Excel
}

Clear-Host
$E = New-Object -ComObject Excel.Application
$E.Visible = $false
$E.DisplayAlerts = $false
$WB = $E.Workbooks.Open($excelFile)
$excelFileName = $WB.Name

EXportWSToCSV -excelFile $excelFile

foreach ($filename in $FileNameList){
    $n = ($filename | split-path -Leaf)
    [System.IO.File]::Copy((("$env:TEMP\" + $n)-replace '"',''),((("$env:TEMP\" + $n + ".txt")-replace '"','')),$true)
    $reader = New-Object System.IO.StreamReader -ArgumentList $filename,([System.Text.Encoding]::Default)
    $writer = New-Object System.IO.StreamWriter -ArgumentList (($filename -replace ".tmp",".txt")),$false,([System.Text.Encoding]::Unicode)
    $writer.AutoFlush = $true
    while (!$reader.EndOfStream){
        $line = $null
        $line = $reader.ReadLine()
        if ($line.trim() -ne ""){
            $writer.WriteLine($line)
        }
    }
    $reader.Close()
    $writer.Close()

    $CSVData = Import-Csv -Path (($filename -replace ".tmp",".txt")) -Encoding Default -Delimiter $tab

    $writer = New-Object System.IO.StreamWriter -ArgumentList (($filename -replace ".tmp",".dat")),$false,([System.Text.Encoding]::Unicode)
    $writer.AutoFlush = $true

    Write-Warning "Please wait while excel is being converted"
    $CSVData | ConvertTo-Csv -NoTypeInformation -Delimiter $pipe | foreach {
        $line = $_ -replace "`n|`r"
        $line = ([regex] '"([^"]*(?:""[^"]*)*)"').Replace($line, { param($match)
            $fieldContent = $match.Groups[1]
            if ($fieldContent -match ("[$pipe"+'"]')) { $match } else { $fieldContent }
        })
        #DEBUGGING write-host $line ; pause ;Clear-Host
        $line = $line -replace ('(?m)"([^'+$pipe+']*?)"(?='+$pipe+'|$)'), '$1'
        #write-host $line ; pause ;Clear-Host
        $line = $line -replace '""','"'
        #write-host $line ; pause ;Clear-Host
        $line = $line -replace $pipe,"$thorne$pipe$thorne"
        #write-host $line ; pause ;Clear-Host
        $line = ("$thorne"+$line+"$thorne")
        #write-host $line ; pause ;Clear-Host
        $writer.WriteLine($line)
    }
$writer.Close()

Clear-Host
Write-Warning "Please wait while loadfile is copied to output"
[System.IO.File]::Copy((($filename -replace ".tmp",".dat")),("$DatLoc\"+((($filename | Split-Path -Leaf) -replace'.tmp','.dat'))),$true)
}
pause
