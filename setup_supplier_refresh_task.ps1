# Дополнительная тихая проверка свежего прайса для остатков Авито.
param(
    [string]$Python = 'C:\Users\TLT-1\AppData\Local\Programs\Python\Python313\pythonw.exe'
)
$ErrorActionPreference = 'Stop'
$taskName = 'БытТехОпт_Сверка_остатков_Авито'
$scriptPath = Join-Path $PSScriptRoot 'refresh_supplier.py'
if (-not (Test-Path -LiteralPath $Python)) { throw 'Python interpreter not found' }
$action = New-ScheduledTaskAction -Execute $Python -Argument ('-X utf8 "' + $scriptPath + '"') -WorkingDirectory $PSScriptRoot
$periodic = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) -RepetitionInterval (New-TimeSpan -Hours 4)
$logon = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 20)
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($periodic, $logon) -Settings $settings -Force | Out-Null
Write-Output ('Registered: ' + $taskName)
