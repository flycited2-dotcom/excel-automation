import json
import sys
from datetime import datetime
from pathlib import Path
from html import escape

BASE_DIR = Path(__file__).parent


def load_config():
    with open(BASE_DIR / 'config.json', encoding='utf-8') as f:
        config = json.load(f)
    # Секреты (токен бота, id канала) держим ВНЕ git — в secrets.json.
    # Если файл есть, его непустые значения перекрывают config.json.
    secrets_file = BASE_DIR / 'secrets.json'
    if secrets_file.exists():
        with open(secrets_file, encoding='utf-8') as f:
            secrets = json.load(f)
        tg = config.setdefault('telegram', {})
        tg.update({k: v for k, v in secrets.get('telegram', {}).items() if v})
    return config


def default_caption(config: dict, note: str = '', issue_date: str = None) -> str:
    """Подпись к ежедневному прайсу — как в канале: эмодзи + активные ссылки."""
    company = escape(config.get('company', {}).get('name', ''))
    return (
        f"✅ <b>Прайс обновлён</b>\n"
        f"{company}\n"
        f"Дата: {issue_date or datetime.today().strftime('%d.%m.%Y')}\n"
        + (f"\n{note}\n" if note else '') +
        f"\n"
        f'<a href="https://splithub.ru/">https://splithub.ru/</a>\n'
        f"Ссылка на приложение:\n"
        f'<a href="https://splithub.ru/app/">https://splithub.ru/app/</a>\n'
        f"\n"
        f"Наш канал в MAX:\n"
        f'<a href="https://max.ru/join/eqkTPKXWFVu-Nz-Os5FlclE1b4EtgdkOQ4QZ0L9PmAE">https://max.ru/join/eqkTPKXWFVu-Nz-Os5FlclE1b4EtgdkOQ4QZ0L9PmAE</a>'
    )


def _document_mime_type(filepath: Path) -> str:
    """MIME-тип для старого .xls и современного .xlsx."""
    if filepath.suffix.lower() == '.xls':
        return 'application/vnd.ms-excel'
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def manager_keyboard(config):
    url = config.get('telegram', {}).get('manager_url', '').strip()
    if not url:
        return None
    from urllib.parse import urlparse
    parsed = urlparse(url)
    if parsed.scheme != 'https' or parsed.hostname != 't.me' or not parsed.path.strip('/'):
        raise ValueError('manager_url must be a Telegram HTTPS contact link')
    return {'inline_keyboard': [[{'text': '💬 Написать менеджеру', 'url': url, 'style': 'success'}]]}


def _send_file(filepath: str, caption: str = None, chat_id: str = None) -> bool:
    """Отправляет Excel-файл в канал или в явно указанный Telegram-чат."""
    try:
        import requests
    except ImportError:
        print("Ошибка: установите пакет requests → pip install requests")
        return False

    config = load_config()
    tg = config.get('telegram', {})
    token = tg.get('bot_token', '').strip()
    if chat_id is None:
        chat_id = tg.get('channel_id', '').strip()
    else:
        chat_id = str(chat_id).strip()

    if not token:
        print("Ошибка: заполните telegram.bot_token в config.json")
        return False
    if not chat_id:
        print("Ошибка: не задан Telegram chat_id для отправки файла")
        return False

    news = None
    if caption is None:
        note, issue_date = '', None
        if str(chat_id) == str(tg.get('channel_id', '')):
            try:
                from daily_price_news import prepare_note
                note, news = prepare_note(filepath, chat_id)
                if news:
                    issue_date = datetime.strptime(news['date'], '%Y-%m-%d').strftime('%d.%m.%Y')
            except Exception as exc:
                print(f"Подпись без аналитики: {type(exc).__name__}")
        caption = default_caption(config, note, issue_date)

    url = f"https://api.telegram.org/bot{token}/sendDocument"
    filepath = Path(filepath)
    data = {'chat_id': chat_id, 'caption': caption, 'parse_mode': 'HTML'}
    if str(chat_id) == str(tg.get('channel_id', '')) and filepath.suffix.lower() == '.xlsx':
        keyboard = manager_keyboard(config)
        if keyboard:
            data['reply_markup'] = json.dumps(keyboard, ensure_ascii=False)

    with open(filepath, 'rb') as f:
        resp = requests.post(
            url,
            files={'document': (filepath.name, f,
                   _document_mime_type(filepath))},
            data=data,
            timeout=60,
        )

    if resp.status_code == 200:
        try:
            from daily_price_news import record_success
            record_success(news, chat_id)
        except Exception as exc:
            print(f"Не удалось сохранить историю подписи: {type(exc).__name__}")
        print(f"Файл отправлен в Telegram: {filepath.name}")
        return True
    else:
        err = resp.json().get('description', resp.text)
        print(f"Ошибка Telegram ({resp.status_code}): {err}")
        return False


def send_file(filepath: str, caption: str = None, chat_id: str = None) -> bool:
    """Доставка учитывается отдельно от успешного формирования итогового прайса."""
    error_type = None
    route = chat_id
    try:
        if route is None:
            route = load_config().get('telegram', {}).get('channel_id', '')
        ok = _send_file(filepath, caption, chat_id)
    except Exception as exc:
        # Текст requests-ошибки может содержать URL с токеном бота.
        error_type = type(exc).__name__
        print(f"Отправка Telegram не выполнена: {error_type}")
        ok = False
    try:
        from ready_price import record_delivery, dispatch
        record_delivery(filepath, 'sent' if ok else 'failed', chat_id=route,
                        error_type=error_type)
        dispatch()
    except Exception as exc:
        print(f"Не удалось записать статус доставки: {type(exc).__name__}")
    return ok


def send_message(text: str, chat_id: str = None) -> bool:
    """Отправляет текстовое сообщение в Telegram.
    По умолчанию — в основной канал; для приватных алертов об ошибках
    передавайте личный chat_id."""
    try:
        import requests
    except ImportError:
        print("Ошибка: установите пакет requests → pip install requests")
        return False

    config = load_config()
    tg = config.get('telegram', {})
    token = tg.get('bot_token', '').strip()
    if not token:
        print("Ошибка: заполните telegram.bot_token в config.json")
        return False
    if chat_id is None:
        chat_id = tg.get('channel_id', '').strip()

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    resp = requests.post(url, data={'chat_id': chat_id, 'text': text}, timeout=30)
    if resp.status_code == 200:
        return True
    err = resp.json().get('description', resp.text)
    print(f"Ошибка Telegram ({resp.status_code}): {err}")
    return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python telegram_send.py <файл.xlsx> [подпись]")
        sys.exit(1)
    caption = sys.argv[2] if len(sys.argv) > 2 else None
    ok = send_file(sys.argv[1], caption)
    sys.exit(0 if ok else 1)
