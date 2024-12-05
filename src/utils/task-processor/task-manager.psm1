# Task management utilities
using namespace System.Threading.Tasks
using namespace System.Collections.Generic

function Wait-ForTasks {
    param (
        [Parameter(Mandatory=$true)]
        [List[Task]]$Tasks,
        
        [Parameter(Mandatory=$false)]
        [timespan]$Timeout = [TimeSpan]::FromMinutes(30)
    )
    
    if ($null -eq $Tasks -or $Tasks.Count -eq 0) { 
        Write-Warning "No tasks to process"
        return $true 
    }
    
    try {
        # Filter out null tasks
        $validTasks = $Tasks.Where({ $null -ne $_ })
        
        if ($validTasks.Count -eq 0) {
            Write-Warning "No valid tasks to process"
            return $true
        }
        
        # Wait for all tasks with timeout
        $completed = [Task]::WaitAll(
            $validTasks.ToArray(),
            $Timeout
        )
        
        if (-not $completed) {
            throw "Task processing timed out after $($Timeout.TotalMinutes) minutes"
        }
        
        # Check for faulted tasks
        $faultedTasks = $validTasks.Where({ $_.IsFaulted })
        if ($faultedTasks.Count -gt 0) {
            $errors = $faultedTasks | ForEach-Object { 
                $_.Exception.InnerExceptions | ForEach-Object { $_.Message }
            }
            throw "Tasks failed with errors: $($errors -join '; ')"
        }
        
        return $true
    }
    catch [OperationCanceledException] {
        Write-Warning "Task processing was cancelled"
        throw
    }
    catch [AggregateException] {
        $_.Exception.InnerExceptions | ForEach-Object {
            Write-Error "Task failed: $($_.Message)"
        }
        throw "One or more tasks failed during parallel processing"
    }
    catch {
        Write-Error "Task processing failed: $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Wait-ForTasks