# For some use cases as of late, this is perfect for me.  But I think over time, will need some tinkering to cover more use cases.
function grep {
    param(
        [Parameter(ValueFromPipeline)]
        [object[]]$InputObject,
        [Parameter(Mandatory, Position = 0)]
        [string]$Pattern
    )
    process {
        try { $InputObject | Format-List -Force | Out-String -Stream | Select-String -Pattern $Pattern }
        catch { Write-Warning "Failed to grep.  Error: $($_.Exception.Message)" }
    }
}
