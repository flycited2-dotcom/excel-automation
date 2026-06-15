import json
import sys
from datetime import datetime
from pathlib import Path

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


def default_caption(config: dict) -> str:
    """Подпись к ежедневному прайсу — как в канале: эмодзи + активные ссылки."""
    company = config.get('company', {}).get('name', '')
    return (
        f"✅ <b>Прайс обновлён</b>\n"
        f"{company}\n"
        f"Дата: {datetime.today().strftime('%d.%m.%Y')}\n"
        f"\n"
        f'<a href="https://splithub.ru/">https://splithub.ru/</a>\n'
        f"Ссылка на приложение:\n"
        f'<a href="https://splithub.ru/app/">https://splithub.ru/app/</a>'
    )


def send_file(filepath: str, caption: str = None) -> bool:
    try:
        import requests
    except ImportError:
        print("Ошибка: установите пакет requests → pip install requests")
        return False

    config = load_config()
    tg = config.get('telegram', {})
    token = tg.get('bot_token', '').strip()
    channel_id = tg.get('channel_id', '').strip()

    if not token:
        print("Ошибка: заполните telegram.bot_token в config.json")
        return False
    if not channel_id:
        print("Ошибка: заполните telegram.channel_id в config.json")
        return False

    if caption is None:
        caption = default_caption(config)

    url = f"https://api.telegram.org/bot{token}/sendDocument"
    filepath = Path(filepath)

    with open(filepath, 'rb') as f:
        resp = requests.post(
            url,
            files={'document': (filepath.name, f,
                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'chat_id': channel_id, 'caption': caption, 'parse_mode': 'HTML'},
            timeout=60,
        )

    if resp.status_code == 200:
        print(f"Файл отправлен в Telegram: {filepath.name}")
        return True
    else:
        err = resp.json().get('description', resp.text)
        print(f"Ошибка Telegram ({resp.status_code}): {err}")
        return False


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
