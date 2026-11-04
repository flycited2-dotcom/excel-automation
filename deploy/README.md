# Работа на VPS

Сервер: `213.109.202.45` (в `~/.ssh/config` — `excel-vps`, ключ `excel_automation_vps`).
Проект: `/opt/excel-automation`, системный пользователь `pricebot`, Python в `.venv`.

## Что запускается

- `excel-price-daily.timer` → `excel-price-daily.service`: ежедневно 09:22 МСК
  `daily.py` (почта ТЛТ → Технолидер → сборка → Telegram), затем `cleanup.sh`.
  `Persistent=true`: если сервер был выключен, прогон стартует после загрузки.
- Avito (`ready_price_sync.py`, `refresh_supplier.py`) на сервере НЕ запускается:
  нет `ready-price-bridge.json`, поэтому `ready_price.dispatch()` ничего не делает.
  Эта цепочка пока остаётся на компьютере TLT-1.

## Обновить код

С компьютера, из корня репозитория:

```bash
git archive HEAD | ssh excel-vps "tar -x -C /opt/excel-automation && chown -R pricebot:pricebot /opt/excel-automation"
```

## Полезные команды на сервере

```bash
systemctl list-timers excel-price-daily.timer      # когда следующий запуск
systemctl start excel-price-daily.service          # запустить прямо сейчас
journalctl -u excel-price-daily.service -n 100     # вывод последних запусков
ls /opt/excel-automation/logs/                     # логи daily_ГГГГММДД.log
```

## Пробный режим

Пока в `secrets.json` на сервере `telegram.channel_id` равен личному чату,
итоговый прайс уходит только владельцу. Для боевого режима это поле удаляется,
и берётся канал из `config.json`.
