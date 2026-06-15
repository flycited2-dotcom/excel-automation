"""Ежедневный автоматический прогон: скачать прайс -> сконвертировать -> отправить.

Запускается Планировщиком Windows без участия человека (см. setup_daily_task.py).
Всё пишется в logs/daily_YYYYMMDD.log. При любой ошибке — запись в лог и короткий
приватный алерт в Telegram (если задан alert_chat_id), чтобы можно было отправить
прайс вручную. В клиентский канал алерты НЕ идут.
"""
import logging
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / 'logs'


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(exist_ok=True)
    logfile = LOG_DIR / f"daily_{datetime.now():%Y%m%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s',
        handlers=[
            logging.FileHandler(logfile, encoding='utf-8'),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger('daily')


def send_alert(log: logging.Logger, text: str) -> None:
    """Шлёт приватный алерт об ошибке, если задан alert_chat_id. Сам никогда не падает."""
    try:
        from download_price import load_secrets
        from telegram_send import send_message
        chat_id = load_secrets().get('telegram', {}).get('alert_chat_id', '')
        if chat_id and not str(chat_id).startswith('ВАШ'):
            send_message(text, str(chat_id))
            log.info("Алерт отправлен в Telegram (%s)", chat_id)
        else:
            log.info("alert_chat_id не задан — алерт только в логе")
    except Exception as e:
        log.warning("Не удалось отправить алерт: %s", e)


def main() -> None:
    # Планировщик Windows запускает скрипт с консолью в cp1251 — без этого
    # print() с кириллицей/эмодзи (напр. ⚠ в transform) роняет весь прогон.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8')
        except Exception:
            pass

    log = setup_logging()
    log.info("=== Ежедневный запуск ===")
    try:
        from download_price import download_price
        from fetch_tlt_email import fetch_tlt_price
        from transform import transform
        from telegram_send import send_file

        log.info("Шаг 0 — забор собственного прайса ТЛТ из почты...")
        try:
            tlt_path = fetch_tlt_price()
            log.info("Прайс ТЛТ обновлён из почты: %s", tlt_path)
        except Exception as e:
            # Не критично: используем последний имеющийся файл в price_TLT_input/.
            log.warning("Не удалось обновить прайс ТЛТ из почты: %s "
                        "(будет использован предыдущий файл, если есть)", e)

        log.info("Шаг 1 — скачивание прайса поставщика...")
        input_path = download_price()
        log.info("Скачано: %s", input_path)

        log.info("Шаг 2 — трансформация прайса...")
        output_path = transform(str(input_path))
        log.info("Готов итоговый прайс: %s", output_path)

        log.info("Шаг 3 — отправка в Telegram...")
        if not send_file(output_path):
            raise RuntimeError("send_file вернул False — отправка не удалась")

        log.info("=== Успех ===")
    except Exception as e:
        log.error("ОШИБКА ежедневного запуска: %s", e, exc_info=True)
        send_alert(
            log,
            f"⚠️ БытТехОпт: автоотправка прайса не удалась.\n"
            f"Причина: {e}\n"
            f"Отправьте прайс вручную."
        )
        sys.exit(1)


if __name__ == '__main__':
    main()
