# Thread calculation utilities
using namespace System

function Get-OptimalThreadCount {
    param (
        [Parameter(Mandatory=$true)]
        [int]$RequestedThreads,
        
        [Parameter(Mandatory=$true)]
        [int]$TotalItems,
        
        [Parameter(Mandatory=$false)]
        [int]$MinThreads = 1,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxSystemThreads = [Environment]::ProcessorCount
    )
    
    try {
        # Validate input parameters
        if ($RequestedThreads -lt 1) {
            Write-Warning "Requested thread count must be at least 1, using minimum"
            $RequestedThreads = $MinThreads
        }
        
        if ($TotalItems -lt 1) {
            Write-Warning "Total items must be at least 1, using minimum"
            $TotalItems = 1
        }
        
        # Calculate system constraints
        $systemMax = [Math]::Max(1, $MaxSystemThreads)
        $effectiveMax = [Math]::Min($RequestedThreads, $systemMax)
        
        # Calculate optimal thread count based on workload
        $optimalThreads = [Math]::Min(
            $effectiveMax,
            [Math]::Max(
                $MinThreads,
                [Math]::Min($TotalItems, [Math]::Floor($TotalItems / 100.0))
            )
        )
        
        return [Math]::Max($MinThreads, $optimalThreads)
    }
    catch {
        Write-Warning "Error calculating optimal thread count: $($_.Exception.Message)"
        return $MinThreads
    }
}

function Get-ChunkSize {
    param (
        [Parameter(Mandatory=$true)]
        [int]$TotalItems,
        
        [Parameter(Mandatory=$true)]
        [int]$ThreadCount,
        
        [Parameter(Mandatory=$false)]
        [int]$MinChunkSize = 1
    )
    
    try {
        # Validate input parameters
        if ($ThreadCount -lt 1) {
            Write-Warning "Thread count must be at least 1, using minimum"
            $ThreadCount = 1
        }
        
        if ($TotalItems -lt 1) {
            Write-Warning "Total items must be at least 1, using minimum"
            $TotalItems = 1
        }
        
        if ($MinChunkSize -lt 1) {
            Write-Warning "Minimum chunk size must be at least 1, using default"
            $MinChunkSize = 1
        }
        
        # Calculate base chunk size
        $baseChunkSize = [Math]::Max(1, [Math]::Ceiling($TotalItems / [double]$ThreadCount))
        
        # Apply minimum chunk size constraint
        $adjustedChunkSize = [Math]::Max($MinChunkSize, $baseChunkSize)
        
        # Ensure chunk size doesn't exceed total items
        return [Math]::Min($adjustedChunkSize, $TotalItems)
    }
    catch {
        Write-Warning "Error calculating chunk size: $($_.Exception.Message)"
        return $MinChunkSize
    }
}

Export-ModuleMember -Function Get-OptimalThreadCount, Get-ChunkSize