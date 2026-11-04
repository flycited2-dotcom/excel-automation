#!/bin/sh
# Удаляет старые рабочие файлы, чтобы не переполнять диск сервера.
# История подписи (output/caption-state) и самые свежие исходники не трогаются.
cd /opt/excel-automation || exit 0
find output -maxdepth 1 -type f -name '*.xlsx' -mtime +30 -delete
find output -maxdepth 1 -type f -name '*.news.json' -mtime +30 -delete
find input -maxdepth 1 -type f \( -name '*.xls' -o -name '*.xlsx' \) -mtime +30 -delete
find input/splithub -type f -mtime +30 -delete 2>/dev/null
find price_TLT_input -maxdepth 1 -type f -name '*.xlsx' -mtime +60 -delete
find ready-price-outbox -mindepth 1 -maxdepth 1 -type d -mtime +30 -exec rm -rf {} + 2>/dev/null
find logs -type f -name '*.log' -mtime +60 -delete
exit 0
