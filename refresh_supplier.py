"""Тихое обновление источника наличия для Авито, без отправки в Telegram."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent


def main() -> int:
    log = BASE / "logs" / "supplier-refresh.log"
    log.parent.mkdir(exist_ok=True)
    try:
        from download_price import download_price
        from transform import transform
        source = download_price(retries=2, delay_sec=15)
        output = transform(str(source), supplier_fallback=False)
        message = f"{datetime.now().isoformat()} accepted_for_sync {Path(output).name}\n"
        result = 0
    except Exception as exc:
        # Исключения сетевых библиотек могут содержать URL с токенами.
        message = f"{datetime.now().isoformat()} error {type(exc).__name__}\n"
        result = 1
    with log.open("a", encoding="utf-8") as stream:
        stream.write(message)
    return result


if __name__ == "__main__":
    raise SystemExit(main())
