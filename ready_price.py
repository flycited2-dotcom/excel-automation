"""Надёжный outbox ИТОГОВОГО прайса, независимый от отправки Telegram.

Снимок фиксируется после успешного сохранения книги. Сбой локального обработчика
не удаляет событие: его повторяет ready_price_sync.py/Планировщик Windows.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
from io import BytesIO
import json
import logging
import os
from pathlib import Path
import re
import subprocess
import sys
import uuid

BASE_DIR = Path(__file__).resolve().parent
SOURCE = "telegram-nikita"
BOT_ID = 8639233666
CHAT_ID = "-1001929222037"


def utcnow():
    return datetime.now(timezone.utc).isoformat()


def atomic(path: Path, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with temp.open("wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def write_json(path: Path, value):
    atomic(path, (json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))


@contextmanager
def locked(root: Path):
    root.mkdir(parents=True, exist_ok=True)
    with (root / ".lock").open("a+b") as stream:
        stream.seek(0)
        stream.write(b"0")
        stream.flush()
        stream.seek(0)
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_LOCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def file_origin(path):
    path = Path(path).resolve()
    stat = path.stat()
    return {"path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "modified_at": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat()}


def capture(output_path, *, input_path, supplier_fallback=False, components=(),
            generated_at=None, evidence="transform_success", outbox=None):
    """Сохранить точные байты готовой книги; повтор SHA не создаёт новое событие."""
    path = Path(output_path).resolve()
    match = re.fullmatch(r"БытТехОпт_(\d{8})(?:_\d{6})?\.xlsx", path.name)
    if not match:
        raise ValueError("Only the finished БытТехОпт workbook is accepted")
    # Не принять первоначальный закупочный прайс даже с переименованным файлом.
    from openpyxl import load_workbook
    data = path.read_bytes()
    book = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        if book.sheetnames != ["Прайс клиента"] or tuple(c.value for c in book.active[2]) != (
            "№", "Артикул", "Бренд", "Наименование", "Цена (руб.)", "Заказ (шт.)"
        ):
            raise ValueError("Unexpected finished workbook structure")
    finally:
        book.close()
    digest = hashlib.sha256(data).hexdigest()
    root = Path(outbox) if outbox is not None else BASE_DIR / "ready-price-outbox"
    event = root / digest
    with locked(root):
        manifest = event / "release.json"
        if manifest.exists():
            if hashlib.sha256((event / "price.xlsx").read_bytes()).hexdigest() != digest:
                raise ValueError("Stored snapshot checksum mismatch")
            return manifest
        origin = file_origin(input_path)
        supplier_data = Path(input_path).read_bytes()
        supplier_file = "supplier" + Path(input_path).suffix.lower()
        if supplier_file not in {"supplier.xls", "supplier.xlsx"}:
            raise ValueError("Unsupported supplier workbook")
        payload = {
            "schema_version": 2, "source": SOURCE, "source_sha256": digest,
            "file": "price.xlsx", "original_filename": path.name,
            "supplier_file": supplier_file,
            "original_path": str(path), "issue_date": datetime.strptime(match[1], "%Y%m%d").date().isoformat(),
            "generated_at": generated_at or utcnow(), "captured_at": utcnow(),
            "provenance": {"producer": "excel-automation.transform", "evidence": evidence,
                           "supplier_fallback": bool(supplier_fallback), "supplier": origin,
                           "components": [file_origin(p) for p in components if p and Path(p).is_file()]},
            "snapshot_kind": "partial" if supplier_fallback else "full",
            "telegram_route": {"bot_id": BOT_ID, "chat_id": CHAT_ID},
        }
        atomic(event / "price.xlsx", data)
        atomic(event / supplier_file, supplier_data)
        write_json(manifest, payload)  # commit marker: пишем ПОСЛЕ книги
    return manifest


def record_delivery(output_path, status, *, chat_id, message_id=None, error_type=None, outbox=None):
    """Обновлять доставку только готового снимка и только в подтверждённый канал."""
    if str(chat_id) != CHAT_ID:
        return
    if status not in {"sent", "failed", "unknown"}:
        raise ValueError("Invalid delivery status")
    path = Path(output_path)
    if not path.is_file():
        return
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    root = Path(outbox) if outbox is not None else BASE_DIR / "ready-price-outbox"
    with locked(root):
        event = root / digest
        if not (event / "release.json").exists():
            return
        delivery = event / "telegram.json"
        payload = json.loads(delivery.read_text(encoding="utf-8")) if delivery.exists() else {"attempts": []}
        attempt = {"status": status, "at": utcnow(), "message_id": message_id, "error_type": error_type}
        payload["attempts"].append(attempt)
        payload["last_attempt"] = attempt
        payload["ever_sent"] = payload.get("ever_sent", False) or status == "sent"
        write_json(delivery, payload)


def dispatch():
    """Не ждать сеть/генератор в процессе конвертации. Повтор — через Планировщик."""
    script = BASE_DIR / "ready_price_sync.py"
    if not script.is_file() or not (BASE_DIR / "ready-price-bridge.json").is_file():
        return
    logs = BASE_DIR / "logs"
    logs.mkdir(exist_ok=True)
    with (logs / "ready-price-sync.log").open("ab") as stream:
        subprocess.Popen([sys.executable, str(script)], cwd=BASE_DIR, stdout=stream, stderr=stream,
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), close_fds=True)


def after_transform(output_path, **kwargs):
    # Ошибка сохранения outbox заметна и должна остановить успешный возврат transform:
    # не допускаем «успешную конвертацию», которая потеряла событие.
    manifest = capture(output_path, **kwargs)
    try:
        dispatch()
    except Exception as exc:
        logging.getLogger(__name__).warning("Ready-price dispatch: %s; snapshot saved", type(exc).__name__)
    return manifest
