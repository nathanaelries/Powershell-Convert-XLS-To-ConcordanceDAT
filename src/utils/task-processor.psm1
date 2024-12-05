# Task processing utilities
using module '.\task-processor\thread-calculator.psm1'
using module '.\task-processor\result-handler.psm1'
using module '.\task-processor\task-manager.psm1'
using module '..\config\settings.psm1'
using namespace System.Collections.Generic
using namespace System.Threading.Tasks

function Start-ParallelProcessing {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$Lines,
        
        [Parameter(Mandatory=$true)]
        [scriptblock]$ProcessingAction,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxThreads = $CONFIG.MAX_THREADS,
        
        [Parameter(Mandatory=$false)]
        [int]$MinChunkSize = 1000
    )
    
    try {
        $totalLines = $Lines.Length

        # Handle empty input case
        if ($totalLines -eq 0) {
            Write-Warning "No lines to process"
            return Convert-ToSortedResults(New-ResultDictionary)
        }

        # Calculate thread and chunk sizes
        $effectiveThreads = Get-OptimalThreadCount -RequestedThreads $MaxThreads -TotalItems $totalLines
        $chunkSize = Get-ChunkSize -TotalItems $totalLines -ThreadCount $effectiveThreads -MinChunkSize $MinChunkSize
        
        # Initialize collections
        $results = New-ResultDictionary
        $tasks = [List[Task]]::new()
        
        # Create and start tasks
        for ($i = 0; $i -lt $totalLines; $i += $chunkSize) {
            $start = $i
            $end = [Math]::Min($start + $chunkSize, $totalLines)
            $currentLines = $Lines[$start..($end-1)]
            
            $taskAction = {
                param($currentLines, $startIndex, $action, $results)
                
                try {
                    foreach ($index in 0..($currentLines.Count-1)) {
                        try {
                            $line = $currentLines[$index]
                            if ([string]::IsNullOrEmpty($line)) { continue }
                            
                            $result = & $action $line
                            [void]$results.TryAdd($($startIndex + $index), $result)
                        }
                        catch {
                            Write-Error "Failed to process line at index $($startIndex + $index): $($_.Exception.Message)"
                        }
                    }
                }
                catch {
                    Write-Error "Task failed at batch starting at index $startIndex: $($_.Exception.Message)"
                    throw
                }
            }
            
            $task = [Task]::Factory.StartNew(
                $taskAction, 
                @($currentLines, $start, $ProcessingAction, $results),
                [TaskCreationOptions]::LongRunning
            )
            
            [void]$tasks.Add($task)
        }
        
        # Wait for tasks to complete
        Wait-ForTasks -Tasks $tasks
        
        # Sort and return results
        $sortedResults = Convert-ToSortedResults -Results $results
        
        if ($sortedResults.Count -eq 0) {
            Write-Warning "No results were produced during processing"
        }
        
        return $sortedResults
    }
    catch {
        throw "Parallel processing failed: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Start-ParallelProcessing