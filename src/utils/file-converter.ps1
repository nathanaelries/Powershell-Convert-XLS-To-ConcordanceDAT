# File conversion utilities
using module '..\config\settings.ps1'
using namespace System.Threading.Tasks
using namespace System.Collections.Concurrent

function Convert-ExcelToTemp {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WorksheetPath,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    $reader = New-Object System.IO.StreamReader -ArgumentList @(
        $WorksheetPath,
        [System.Text.Encoding]::Default,
        $true,
        $CONFIG.BUFFER_SIZE
    )
    
    $writer = New-Object System.IO.StreamWriter -ArgumentList @(
        $OutputPath,
        $false,
        [System.Text.Encoding]::Unicode,
        $CONFIG.BUFFER_SIZE
    )
    
    $writer.AutoFlush = $false
    $buffer = New-Object char[] $CONFIG.READ_WRITE_BUFFER
    $bytesRead = 0

    while (($bytesRead = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
        $writer.Write($buffer, 0, $bytesRead)
    }

    $writer.Flush()
    $reader.Close()
    $writer.Close()
}

function Convert-TempToDAT {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    $lines = [System.IO.File]::ReadLines($InputPath)
    $batchSize = $CONFIG.BATCH_SIZE
    $queue = [ConcurrentQueue[string]]::new()
    $writer = [System.IO.StreamWriter]::new($OutputPath, $false, [System.Text.Encoding]::Unicode, $CONFIG.BUFFER_SIZE)
    
    try {
        $batch = [System.Collections.Generic.List[string]]::new($batchSize)
        
        foreach ($line in $lines) {
            $batch.Add($line)
            
            if ($batch.Count -ge $batchSize) {
                Process-Batch $batch $writer
                $batch.Clear()
            }
        }
        
        # Process remaining lines
        if ($batch.Count -gt 0) {
            Process-Batch $batch $writer
        }
    }
    finally {
        $writer.Flush()
        $writer.Close()
    }
}

function Process-Batch {
    param (
        [System.Collections.Generic.List[string]]$Batch,
        [System.IO.StreamWriter]$Writer
    )
    
    $tasks = @()
    $batchArray = $Batch.ToArray()
    $chunkSize = [Math]::Ceiling($batchArray.Count / $CONFIG.MAX_THREADS)
    
    for ($i = 0; $i -lt $CONFIG.MAX_THREADS; $i++) {
        $start = $i * $chunkSize
        $end = [Math]::Min($start + $chunkSize, $batchArray.Count)
        
        if ($start -lt $batchArray.Count) {
            $task = [Task]::Run({
                param($lines, $start, $end)
                $result = @()
                for ($j = $start; $j -lt $end; $j++) {
                    $result += Format-DATLine $lines[$j]
                }
                return $result
            }.GetNewClosure()).AsTask()
            
            $tasks += @{
                Task = $task
                Start = $start
                End = $end
            }
        }
    }
    
    foreach ($task in $tasks) {
        $results = $task.Task.Result
        foreach ($line in $results) {
            $Writer.WriteLine($line)
        }
    }
}

function Format-DATLine {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Line
    )

    $line = $Line -replace "`n|`r"
    
    # Optimize regex operations by pre-compiling patterns
    $static:regexQuoted = [regex]::new('"([^"]*(?:""[^"]*)*)"', 
        [System.Text.RegularExpressions.RegexOptions]::Compiled)
    
    $line = $static:regexQuoted.Replace($line, { 
        param($match)
        $fieldContent = $match.Groups[1]
        if ($fieldContent -match ("[$($CONFIG.PIPE)"+'"]')) { 
            $match 
        } else { 
            $fieldContent 
        }
    })
    
    $line = $line -replace ('(?m)"([^'+$CONFIG.PIPE+']*?)"(?='+$CONFIG.PIPE+'|$)'), '$1'
    $line = $line -replace '""','"'
    $line = $line -replace $CONFIG.PIPE,"$($CONFIG.THORNE)$($CONFIG.PIPE)$($CONFIG.THORNE)"
    $line = "$($CONFIG.THORNE)${line}$($CONFIG.THORNE)"
    
    return $line
}

Export-ModuleMember -Function Convert-ExcelToTemp, Convert-TempToDAT