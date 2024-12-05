# Cleanup utilities
function Remove-TempFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$FilePaths
    )
    
    foreach ($file in $FilePaths) {
        if (Test-Path $file) {
            try {
                Remove-Item -Path $file -Force -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to remove temporary file: $file"
            }
        }
    }
}

Export-ModuleMember -Function Remove-TempFiles