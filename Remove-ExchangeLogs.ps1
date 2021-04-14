#Requires -Version 4

$MaxLogFileAge = 30
$MaxCleanupLogFileAge = 90

$DateTime = [DateTime]::Now

$ExLogFolderPaths = @(

    'C:\inetpub\logs\LogFiles\',
    'C:\Program Files\Microsoft\Exchange Server\V15\Logging\',
    'C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\ETLTraces\',
    'C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\Logs'
)


$LogFolder = Join-Path -Path $PSScriptRoot -ChildPath 'Logging_Remove-ExchangeLogs.ps1'
if (-not (Test-Path -Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory }

function Write-Log ([string]$Message) {

    $LogFile = Join-Path -Path $LogFolder -ChildPath "Exchange-Log-Cleanup_$($DateTime.ToString('yyyy-MM-dd')).log"
    if (-not (Test-Path -Path $LogFile)) { New-Item -Path $LogFile -ItemType:File }

    $MessageData = "[$($DateTime.ToString('yyyy-MM-dd hh:mm:ss tt (zzzz)'))] $($Message)"

    if ($PSVersionTable.PSVersion.Major -eq 4) { Write-Host -Object $MessageData }
    else { Write-Information -MessageData $MessageData -InformationAction Continue }

    $MessageData | Out-File -FilePath $LogFile -Append
}

function Remove-ExchangeLogs ([System.IO.FileInfo[]]$ExLogFolderPaths, [int]$MaxLogFileAge) {

    Write-Log -Message "Remove-ExchangeLog.ps1 starting: MaxLogFileAge = $($MaxLogFileAge), Folders to process:`r`n`t$($ExLogFolderPaths -join ""`r`n`t"")"

    foreach ($Folder in $ExLogFolderPaths) {

        Write-Log -Message "Processing folder: $($Folder)"

        if (Test-Path $Folder) {

            $FilesToDelete = Get-ChildItem -Path $Folder -Recurse -Include *.log, *.blg, *.etl |
            Where-Object {$_.LastWriteTime -le $DateTime.AddDays(-$MaxLogFileAge)}

            foreach ($File in $FilesToDelete) {

                try {
                    Remove-Item $File.FullName -ErrorAction Stop | Out-Null
                    Write-Log -Message "Successfully deleted '$($File.FullName)'."
                }
                catch {
                    Write-Log -Message "Failed to delete '$($File.FullName)'.`r`n`tException:`r`n$($_.Exception)"
                }
            }
        }
        else {
            Write-Log -Message "Folder '$($Folder)' was not found."
        }
    }

    Write-Log -Message 'Remove-ExchangeLog.ps1 finished/ending.'
}

function Remove-CleanupLogs ([int]$MaxCleanupLogFileAge) {

    Write-Log -Message "Remove-CleanupLog.ps1 starting: MaxCleanupLogFileAge = $($MaxCleanupLogFileAge)"

    $CleanupLogsToDelete = Get-ChildItem -Path $Folder -Recurse -Include 'Exchange-Log-Cleanup_*.log' |
    Where-Object {$_.LastWriteTime -le $DateTime.AddDays(-$MaxCleanupLogFileAge)}

    foreach ($CleanupLog in $CleanupLogsToDelete) {

        try {
            Remove-Item $CleanupLog.FullName -ErrorAction Stop | Out-Null
            Write-Log -Message "[Housekeeping] Successfully deleted '$($CleanupLog.FullName)'."
        }
        catch {
            Write-Log -Message "[Housekeeping] Failed to delete '$($CleanupLog.FullName)'.`r`n`tException:`r`n$($_.Exception)"
        }
    }

    Write-Log -Message 'Remove-CleanupLog.ps1 finished/ending.'
}

Remove-ExchangeLogs -ExLogFolderPaths $ExLogFolderPaths -MaxLogFileAge $MaxLogFileAge
Remove-CleanupLogs -MaxCleanupLogFileAge $MaxCleanupLogFileAge
