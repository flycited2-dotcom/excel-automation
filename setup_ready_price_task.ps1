# Повтор обработки сохранённых выпусков после сбоя процесса или перезапуска ПК.
# Задача не отправляет сообщения и не публикует объявления.
param(
    [string]$Python = 'C:\Users\TLT-1\AppData\Local\Programs\Python\Python313\pythonw.exe'
)
$ErrorActionPreference = 'Stop'
$taskName = 'БытТехОпт_Импорт_Авито'
$scriptPath = Join-Path $PSScriptRoot 'ready_price_sync.py'
if (-not (Test-Path -LiteralPath $Python)) { throw 'Python interpreter not found' }
if (-not (Test-Path -LiteralPath $scriptPath)) { throw 'Ready-price sync script not found' }
$action = New-ScheduledTaskAction -Execute $Python -Argument ('-X utf8 "' + $scriptPath + '"') -WorkingDirectory $PSScriptRoot
$periodic = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 15)
$logon = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($periodic, $logon) -Settings $settings -Force | Out-Null
Write-Output ('Registered: ' + $taskName)
