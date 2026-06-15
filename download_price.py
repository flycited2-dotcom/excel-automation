"""Скачивание прайса поставщика (technoleader.net) в папку input/.

Сайт поставщика часто недоступен, поэтому скачивание делает несколько попыток.
Скачанный файл проверяется на то, что это настоящий Excel, а не HTML-страница
ошибки/разлогина. Конкретные шаги входа на сайт — в _fetch_price_file().
"""
import json
import shutil
import time
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
SECRETS_FILE = BASE_DIR / 'secrets.json'
INPUT_DIR = BASE_DIR / 'input'


def load_secrets() -> dict:
    if not SECRETS_FILE.exists():
        raise FileNotFoundError(
            f"Нет файла секретов: {SECRETS_FILE}\n"
            f"Скопируйте secrets.example.json в secrets.json и впишите логин/пароль поставщика."
        )
    with open(SECRETS_FILE, encoding='utf-8') as f:
        return json.load(f)


def _looks_like_excel(path: Path) -> bool:
    """Грубая проверка: это файл Excel, а не HTML-страница ошибки сайта.
    .xls (OLE2) начинается с D0CF11E0, .xlsx (zip) — с 'PK\\x03\\x04'."""
    if not path.exists() or path.stat().st_size < 2048:
        return False
    head = path.read_bytes()[:4]
    return head == b'\xD0\xCF\x11\xE0' or head == b'PK\x03\x04'


def _fetch_price_file(dest: Path, supplier: dict) -> None:
    """Скачать прайс с technoleader.net (WordPress + плагин Download Manager) в dest.

    Два пароля — как при ручном скачивании:
      1) post_password снимает защиту страницы партнёров (ставит cookie);
      2) пароль пакета WPDM: POST с execute=wpdm_getlink сразу отдаёт сам .xls.
    Если пароль неверный / сайт лежит — вернётся не Excel, и download_price() это отбракует.
    """
    import requests

    base = supplier.get('url', 'https://technoleader.net/').rstrip('/') + '/'
    package_id = str(supplier.get('package_id', 658))
    password = str(supplier.get('password', ''))
    if not password:
        raise ValueError("Не задан supplier.password в secrets.json")

    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})

    # 1) пароль страницы (штатная защита WordPress) -> cookie wp-postpass_*
    session.post(base + 'wp-login.php?action=postpass',
                 data={'post_password': password, 'Submit': 'Enter'}, timeout=20)

    # 2) пароль пакета WPDM -> тело ответа это сам файл прайса (.xls)
    ctz = int(time.time() * 1000) % 1000
    resp = session.post(f"{base}?nocache={ctz}", data={
        '__wpdm_ID': package_id,
        'dataType': 'json',
        'execute': 'wpdm_getlink',
        'action': 'wpdm_ajax_call',
        'password': password,
    }, timeout=60)
    resp.raise_for_status()
    dest.write_bytes(resp.content)


def download_price(retries: int = 3, delay_sec: int = 30) -> Path:
    """Скачивает прайс поставщика в input/ и возвращает путь к файлу.

    Делает до `retries` попыток с паузой `delay_sec` секунд между ними —
    сайт поставщика часто недоступен.
    """
    secrets = load_secrets()
    supplier = secrets.get('supplier', {})
    INPUT_DIR.mkdir(exist_ok=True)

    dest = INPUT_DIR / f"priceopt_{datetime.now():%Y%m%d}.xls"
    tmp = INPUT_DIR / '.download_tmp'

    last_err = None
    for attempt in range(1, retries + 1):
        try:
            if tmp.exists():
                tmp.unlink()
            _fetch_price_file(tmp, supplier)
            if not _looks_like_excel(tmp):
                raise ValueError("Скачан не Excel-файл (возможно, страница ошибки сайта)")
            shutil.move(str(tmp), str(dest))
            return dest
        except NotImplementedError:
            raise  # повторять стуб бессмысленно
        except Exception as e:
            last_err = e
            print(f"Попытка {attempt}/{retries} не удалась: {e}")
            if attempt < retries:
                time.sleep(delay_sec)

    raise RuntimeError(f"Не удалось скачать прайс за {retries} попыток: {last_err}")


if __name__ == '__main__':
    path = download_price()
    print(f"Скачано: {path}")
