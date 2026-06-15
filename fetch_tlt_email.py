"""Забор собственного прайса (ТЛТ) из почты в папку price_TLT_input/.

1С автоматически шлёт внутренний оптовый прайс письмом (отправитель
info@simfer.com.ru, тема «Прайс_Опт_внутрн ...», в теле — «Для скрипта»).
Этот модуль по IMAP находит самое свежее такое письмо, скачивает
вложение .xlsx и кладёт в price_TLT_input/. Дальше штатно: pick_tlt_file()
берёт самый свежий файл, read_tlt_price() его парсит.

Авторизация — Gmail IMAP по паролю приложения (нужна включённая 2FA на
аккаунте). Логин/пароль — в secrets.json, блок "email". См. secrets.example.json.
"""
import email
import imaplib
import json
from datetime import datetime
from email.header import decode_header
from pathlib import Path

BASE_DIR = Path(__file__).parent
SECRETS_FILE = BASE_DIR / 'secrets.json'
TLT_DIR = BASE_DIR / 'price_TLT_input'

# Значения по умолчанию — перекрываются блоком "email" в secrets.json.
DEFAULTS = {
    'imap_host': 'imap.gmail.com',
    'imap_port': 993,
    'sender': 'info@simfer.com.ru',
    'subject_contains': 'Прайс_Опт_авторассылка',
}


def _load_email_secrets() -> dict:
    if not SECRETS_FILE.exists():
        raise FileNotFoundError(
            f"Нет файла секретов: {SECRETS_FILE}\n"
            f"Скопируйте secrets.example.json в secrets.json и заполните блок \"email\"."
        )
    with open(SECRETS_FILE, encoding='utf-8') as f:
        secrets = json.load(f)
    cfg = dict(DEFAULTS)
    cfg.update({k: v for k, v in secrets.get('email', {}).items() if v != ''})
    if not cfg.get('user') or not cfg.get('app_password'):
        raise ValueError(
            "В secrets.json не заданы email.user / email.app_password.\n"
            "user — адрес ящика (flycited@gmail.com), app_password — пароль приложения Google."
        )
    return cfg


def _decode(value) -> str:
    """Декодирует MIME-заголовок (тема/имя файла) в обычную строку."""
    if value is None:
        return ''
    parts = []
    for text, enc in decode_header(value):
        if isinstance(text, bytes):
            parts.append(text.decode(enc or 'utf-8', errors='replace'))
        else:
            parts.append(text)
    return ''.join(parts)


def _looks_like_excel(data: bytes) -> bool:
    """Грубая проверка: это Excel (.xls OLE2 или .xlsx zip), а не мусор/HTML."""
    if not data or len(data) < 2048:
        return False
    head = data[:4]
    return head == b'\xD0\xCF\x11\xE0' or head == b'PK\x03\x04'


def _pick_xlsx_attachment(msg) -> tuple[str, bytes] | None:
    """Возвращает (имя_файла, содержимое) первого вложения .xlsx/.xls.
    PDF и прочее игнорируются."""
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        filename = _decode(part.get_filename())
        if not filename:
            continue
        low = filename.lower()
        if low.endswith('.xlsx') or low.endswith('.xls'):
            payload = part.get_payload(decode=True)
            if payload:
                return filename, payload
    return None


def fetch_tlt_price() -> Path:
    """Скачивает самый свежий прайс ТЛТ из почты в price_TLT_input/.

    Возвращает путь к сохранённому файлу. Бросает исключение, если письмо
    не найдено или вложение оказалось не Excel.
    """
    cfg = _load_email_secrets()
    TLT_DIR.mkdir(exist_ok=True)

    imap = imaplib.IMAP4_SSL(cfg['imap_host'], int(cfg['imap_port']))
    try:
        imap.login(cfg['user'], cfg['app_password'])
        imap.select('INBOX', readonly=True)

        # Ищем по отправителю (надёжнее, чем по кириллической теме в IMAP SEARCH).
        # Тему проверяем уже в Python после декодирования заголовка.
        typ, data = imap.uid('search', None, 'FROM', f'"{cfg["sender"]}"')
        if typ != 'OK' or not data or not data[0]:
            raise RuntimeError(
                f"Не найдено писем от {cfg['sender']}. "
                f"Проверьте, что 1С прислала прайс и письмо во входящих."
            )

        # UID растут со временем — берём от новых к старым, останавливаемся
        # на первом письме, чья тема содержит нужную подстроку.
        uids = data[0].split()
        subject_needle = cfg['subject_contains'].lower()
        for uid in reversed(uids):
            typ, msg_data = imap.uid('fetch', uid, '(RFC822)')
            if typ != 'OK' or not msg_data or msg_data[0] is None:
                continue
            msg = email.message_from_bytes(msg_data[0][1])
            subject = _decode(msg.get('Subject'))
            if subject_needle not in subject.lower():
                continue

            attach = _pick_xlsx_attachment(msg)
            if attach is None:
                raise RuntimeError(
                    f"В письме «{subject}» нет вложения .xlsx (есть только PDF?)."
                )
            filename, content = attach
            if not _looks_like_excel(content):
                raise ValueError(
                    f"Вложение «{filename}» не похоже на Excel-файл."
                )

            dest = TLT_DIR / f"Прайс_ТЛТ_{datetime.now():%Y%m%d}.xlsx"
            dest.write_bytes(content)
            return dest

        raise RuntimeError(
            f"Письма от {cfg['sender']} есть, но без темы «{cfg['subject_contains']}». "
            f"Проверьте тему рассылки из 1С."
        )
    finally:
        try:
            imap.logout()
        except Exception:
            pass


if __name__ == '__main__':
    path = fetch_tlt_price()
    print(f"Скачано: {path}")
