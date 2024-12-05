# Progress reporting utilities
using module '..\config\settings.psm1'

function Write-ProcessProgress {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Activity,
        
        [Parameter(Mandatory=$true)]
        [string]$Status,
        
        [Parameter(Mandatory=$true)]
        [int]$Current,
        
        [Parameter(Mandatory=$true)]
        [int]$Total
    )
    
    $percentComplete = [math]::Round(($Current / [Math]::Max(1, $Total)) * 100)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}

Export-ModuleMember -Function Write-ProcessProgress