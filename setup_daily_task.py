"""Регистрация ЕЖЕДНЕВНОЙ задачи Windows: каждый день в заданное время запускает daily.py
(скачать прайс -> сконвертировать -> отправить в Telegram).

Время берётся из config.json -> schedule.daily_time (по умолчанию 09:00).
Существующий разовый schedule_task.py не затрагивается.
"""
import json
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
TASK_NAME = 'БытТехОпт_Ежедневный_прайс'


def main() -> None:
    with open(BASE_DIR / 'config.json', encoding='utf-8') as f:
        config = json.load(f)
    daily_time = config.get('schedule', {}).get('daily_time', '09:00')

    try:
        hh, mm = map(int, daily_time.split(':'))
        assert 0 <= hh <= 23 and 0 <= mm <= 59
    except (ValueError, AssertionError):
        print(f"Неверное время в config.json -> schedule.daily_time: {daily_time!r}. Нужен формат HH:MM")
        sys.exit(1)

    python_exe = sys.executable
    daily_script = str(BASE_DIR / 'daily.py')

    # PowerShell-регистрация (не зависит от локали Windows), как в schedule_task.py
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$action  = New-ScheduledTaskAction -Execute '{python_exe}' -Argument '"{daily_script}"' -WorkingDirectory '{BASE_DIR}'
$trigger = New-ScheduledTaskTrigger -Daily -At '{hh:02d}:{mm:02d}'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable
Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host 'OK'
"""

    result = subprocess.run(
        ['powershell', '-NoProfile', '-Command', ps_script],
        capture_output=True, text=True
    )

    if 'OK' in result.stdout and not result.stderr.strip():
        print()
        print(f"  Готово! Ежедневная задача '{TASK_NAME}' создана.")
        print(f"  Запуск каждый день в {hh:02d}:{mm:02d}.")
        print(f"  Сработает, когда компьютер включён и выполнен вход в систему.")
        print(f"  (если ПК был выключен в это время — запустится при следующем включении)")
        print()
    else:
        print(f"  Ошибка создания задачи:")
        print(result.stderr or result.stdout or "(нет вывода)")
        sys.exit(1)


if __name__ == '__main__':
    main()
