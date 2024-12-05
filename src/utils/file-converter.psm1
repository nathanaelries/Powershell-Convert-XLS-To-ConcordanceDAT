# File conversion utilities
using module '..\config\settings.psm1'

function Convert-ExcelToTemp {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WorksheetPath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    try {
        if (-not (Test-Path $WorksheetPath)) {
            throw "Input worksheet file not found: $WorksheetPath"
        }

        $reader = [System.IO.StreamReader]::new(
            $WorksheetPath,
            [System.Text.Encoding]::Default,
            $true,
            $CONFIG.BUFFER_SIZE
        )
        
        $writer = [System.IO.StreamWriter]::new(
            $OutputPath,
            $false,
            [System.Text.Encoding]::Unicode,
            $CONFIG.BUFFER_SIZE
        )
        
        $writer.AutoFlush = $false
        $buffer = New-Object char[] $CONFIG.READ_WRITE_BUFFER
        
        while (($bytesRead = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $writer.Write($buffer, 0, $bytesRead)
        }

        $writer.Flush()
    }
    catch {
        throw "Failed to convert Excel to temp: $($_.Exception.Message)"
    }
    finally {
        if ($writer) { 
            try { $writer.Close() } catch { }
            try { $writer.Dispose() } catch { }
        }
        if ($reader) { 
            try { $reader.Close() } catch { }
            try { $reader.Dispose() } catch { }
        }
    }
}

function Convert-TempToDAT {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    try {
        if (-not (Test-Path $InputPath)) {
            throw "Input file not found: $InputPath"
        }

        # Create output directory if it doesn't exist
        $outputDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }

        $reader = [System.IO.StreamReader]::new(
            $InputPath,
            [System.Text.Encoding]::Default,
            $true,
            $CONFIG.BUFFER_SIZE
        )

        $writer = [System.IO.StreamWriter]::new(
            $OutputPath,
            $false,
            [System.Text.Encoding]::Unicode,
            $CONFIG.BUFFER_SIZE
        )

        $writer.AutoFlush = $false
        $lineCount = 0
        $batchCount = 0

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $formattedLine = Format-DATLine $line
                if ($formattedLine) {
                    $writer.WriteLine($formattedLine)
                    $lineCount++
                }
            }

            # Flush periodically to manage memory
            if ($lineCount % $CONFIG.BATCH_SIZE -eq 0) {
                $writer.Flush()
                $batchCount++
                Write-Progress -Activity "Converting to DAT" -Status "Processed $lineCount lines" -PercentComplete -1
            }
        }

        # Final flush
        $writer.Flush()
        Write-Progress -Activity "Converting to DAT" -Status "Completed processing $lineCount lines" -Completed
    }
    catch {
        throw "Failed to convert temp to DAT: $($_.Exception.Message)"
    }
    finally {
        if ($writer) { 
            try { $writer.Close() } catch { }
            try { $writer.Dispose() } catch { }
        }
        if ($reader) {
            try { $reader.Close() } catch { }
            try { $reader.Dispose() } catch { }
        }
    }
}

function Format-DATLine {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Line
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Line)) {
            return $null
        }

        $line = $Line -replace "`n|`r"
        
        if (-not $script:regexQuoted) {
            $script:regexQuoted = [regex]::new(
                '"([^"]*(?:""[^"]*)*)"', 
                [System.Text.RegularExpressions.RegexOptions]::Compiled
            )
        }
        
        $line = $script:regexQuoted.Replace($line, { 
            param($match)
            $fieldContent = $match.Groups[1].Value
            if ($fieldContent -match [regex]::Escape("$($CONFIG.PIPE)") -or $fieldContent -match '"') { 
                $match.Value
            } else { 
                $fieldContent 
            }
        })
        
        $line = $line -replace ('(?m)"([^'+[regex]::Escape("$($CONFIG.PIPE)")+']*?)"(?='+[regex]::Escape("$($CONFIG.PIPE)")+'|$)'), '$1'
        $line = $line -replace '""','"'
        $line = $line -replace [regex]::Escape("$($CONFIG.PIPE)"),"$($CONFIG.THORNE)$($CONFIG.PIPE)$($CONFIG.THORNE)"
        $line = "$($CONFIG.THORNE)${line}$($CONFIG.THORNE)"
        
        return $line
    }
    catch {
        throw "Failed to format DAT line: $($_.Exception.Message)"
    }
}

# Create module-level variables
$script:regexQuoted = $null

Export-ModuleMember -Function Convert-ExcelToTemp, Convert-TempToDAT, Format-DATLine