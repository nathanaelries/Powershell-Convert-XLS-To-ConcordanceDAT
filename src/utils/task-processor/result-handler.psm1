# Result handling utilities
using namespace System.Collections.Generic
using namespace System.Collections.Concurrent

function New-ResultDictionary {
    return [ConcurrentDictionary[int,string]]::new()
}

function Convert-ToSortedResults {
    param (
        [Parameter(Mandatory=$true)]
        [ConcurrentDictionary[int,string]]$Results
    )
    
    $sortedResults = [SortedDictionary[int,string]]::new()
    
    foreach ($kvp in $Results.GetEnumerator()) {
        $sortedResults[$kvp.Key] = $kvp.Value
    }
    
    return $sortedResults
}

Export-ModuleMember -Function New-ResultDictionary, Convert-ToSortedResults