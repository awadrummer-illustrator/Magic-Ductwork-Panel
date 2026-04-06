$ErrorActionPreference = 'Stop'

$taskName = 'Reload Illustrator Ductwork'
$buildPlugin = 'E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Process Ductwork\build\win\x64\Release\ProcessDuctwork.aip'
$installedPlugin = 'C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\ProcessDuctwork.aip'

if (-not (Test-Path -LiteralPath $buildPlugin)) {
	throw "Built plugin not found: $buildPlugin"
}

$task = Get-ScheduledTask -TaskName $taskName
$beforeInfo = Get-ScheduledTaskInfo -TaskName $taskName
$beforeRunTime = $beforeInfo.LastRunTime

Write-Host "Running scheduled task: schtasks /run /tn `"$taskName`""
schtasks /run /tn $taskName | Out-Host

$deadline = (Get-Date).AddMinutes(2)
do {
	Start-Sleep -Seconds 2
	$task = Get-ScheduledTask -TaskName $taskName
	$afterInfo = Get-ScheduledTaskInfo -TaskName $taskName
	$ranThisTime = $afterInfo.LastRunTime -gt $beforeRunTime
	$isStillRunning = $task.State -eq 'Running' -or $task.State -eq 'Queued'
} while ((-not $ranThisTime -or $isStillRunning) -and (Get-Date) -lt $deadline)

if (-not $ranThisTime) {
	throw "Scheduled task did not report a new run within the timeout."
}

if ($afterInfo.LastTaskResult -ne 0) {
	throw "Scheduled task finished with LastTaskResult=$($afterInfo.LastTaskResult)."
}

if (-not (Test-Path -LiteralPath $installedPlugin)) {
	throw "Installed plugin not found: $installedPlugin"
}

$buildHash = Get-FileHash -LiteralPath $buildPlugin -Algorithm SHA256
$installedHash = Get-FileHash -LiteralPath $installedPlugin -Algorithm SHA256
if ($buildHash.Hash -ne $installedHash.Hash) {
	throw "Installed plugin hash does not match build output.`nBuilt:     $($buildHash.Hash)`nInstalled: $($installedHash.Hash)"
}

$installedItem = Get-Item -LiteralPath $installedPlugin
$illustrator = Get-Process Illustrator -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1

Write-Host ''
Write-Host 'Verification passed.'
Write-Host "Task: $taskName"
Write-Host "Last Run Time: $($afterInfo.LastRunTime)"
Write-Host "Last Result: $($afterInfo.LastTaskResult)"
Write-Host "Installed Plugin: $installedPlugin"
Write-Host "Installed Plugin LastWriteTime: $($installedItem.LastWriteTime)"
Write-Host "Installed Plugin Hash: $($installedHash.Hash)"
Write-Host "Built Plugin Hash: $($buildHash.Hash)"
if ($illustrator) {
	Write-Host "Illustrator PID: $($illustrator.Id)"
	Write-Host "Illustrator StartTime: $($illustrator.StartTime)"
} else {
	Write-Warning 'Illustrator is not currently running.'
}
