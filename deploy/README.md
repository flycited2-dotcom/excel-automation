# Работа на VPS

Сервер: `213.109.202.45` (в `~/.ssh/config` — `excel-vps`, ключ `excel_automation_vps`).
Проект: `/opt/excel-automation`, системный пользователь `pricebot`, Python в `.venv`.

## Что запускается

- `excel-price-daily.timer` → `excel-price-daily.service`: ежедневно 09:22 МСК
  `daily.py` (почта ТЛТ → Технолидер → сборка → Telegram), затем доставка в Avito
  и `cleanup.sh`. `Persistent=true`: если сервер был выключен, прогон стартует после загрузки.
- `excel-supplier-refresh.timer` — каждые 4 ч: `refresh_supplier.py` (свежий прайс
  Технолидера → сборка без Telegram) и доставка выпуска в Avito.
- `excel-avito-sync.timer` — каждые 15 мин: `ready_price_sync.py` повторяет
  недоставленные выпуски.

## Avito

`ready-price-bridge.json` на сервере = `deploy/ready-price-bridge.vps.json`,
режим `push_only`: локальный avito-bridge не нужен. Полные выпуски (не fallback)
архивом уходят в `/opt/avito-ready-price/inbox` на Sprintbox (212.116.115.150:2222),
проверяет и публикует их `avito-ready-price.timer` на Sprintbox.

Доступ: ключ `/opt/excel-automation/.ssh/id_avito` (хост `sprintbox-avito` в
`/opt/excel-automation/.ssh/config`). На Sprintbox он ограничен в `authorized_keys`
командой `/usr/local/bin/receive-ready-price` (`deploy/receive-ready-price`):
принимает только архив в inbox, консоль и чтение недоступны.

Отметка доставки: `ready-price-outbox/<sha>/remote-push.json`; итог прохода —
`ready-price-outbox/sync-status.json`.

На ПК TLT-1 с 04.11.2026 доставка в Avito выключена (`remote_push.enabled=false`
в локальном `ready-price-bridge.json`, копия — `.bak-before-vps`), задачи
`БытТехОпт_Импорт_Авито` и `БытТехОпт_Сверка_остатков_Авито` отключены.

## Обновить код

С компьютера, из корня репозитория:

```bash
git archive HEAD | ssh excel-vps "tar -x -C /opt/excel-automation && chown -R pricebot:pricebot /opt/excel-automation"
```

## Полезные команды на сервере

```bash
systemctl list-timers "excel-*"                    # когда следующие запуски
systemctl start excel-price-daily.service          # запустить прямо сейчас
journalctl -u excel-price-daily.service -n 100     # вывод последних запусков
ls /opt/excel-automation/logs/                     # логи daily_ГГГГММДД.log
```

## Боевой режим (с 06.11.2026)

Сервер шлёт прайс в клиентский канал из `config.json`; задача
`БытТехОпт_Ежедневный_прайс` на ПК TLT-1 выключена (не удалена).
Копия пробных настроек — `secrets.json.bak-trial` (там `telegram.channel_id`
= личный чат). Вернуть пробный режим: `cp -p secrets.json.bak-trial secrets.json`.

С VPS Telegram доступен только по IPv6 (IPv4 до api.telegram.org не проходит),
и IPv6 изредка не отвечает. Поэтому `telegram_send.send_file` повторяет
отправку до 4 раз, но только при ошибке подключения — без риска дублей.
