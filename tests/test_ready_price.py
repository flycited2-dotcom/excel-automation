from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys

from openpyxl import Workbook
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ready_price
import telegram_send


def workbook(tmp_path):
    path = tmp_path / f"БытТехОпт_{datetime.now():%Y%m%d}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Прайс клиента"
    ws.append([None] * 6)
    ws.append(["№", "Артикул", "Бренд", "Наименование", "Цена (руб.)", "Заказ (шт.)"])
    ws.append(["Мультиварки"])
    ws.append([1, "000123", "Comfee", "CF-MC9521", 3993, 99])
    wb.save(path)
    supplier = tmp_path / "supplier.xls"
    supplier.write_bytes(b"raw supplier, not the finished workbook")
    return path, supplier


def test_capture_is_immutable_and_idempotent(tmp_path):
    output, source = workbook(tmp_path)
    outbox = tmp_path / "outbox"
    first = ready_price.capture(output, input_path=source, outbox=outbox)
    before = first.read_bytes()
    second = ready_price.capture(output, input_path=source, outbox=outbox)
    assert first == second and first.read_bytes() == before
    data = json.loads(before)
    assert (first.parent / "price.xlsx").read_bytes() == output.read_bytes()
    assert data["source_sha256"] == hashlib.sha256(output.read_bytes()).hexdigest()
    assert data["provenance"]["supplier"]["sha256"] == hashlib.sha256(source.read_bytes()).hexdigest()
    assert not (first.parent / "telegram.json").exists()


def test_delivery_failure_does_not_lose_snapshot(tmp_path):
    output, source = workbook(tmp_path)
    outbox = tmp_path / "outbox"
    manifest = ready_price.capture(output, input_path=source, outbox=outbox)
    before = manifest.read_bytes()
    ready_price.record_delivery(output, "failed", chat_id=ready_price.CHAT_ID, error_type="Timeout", outbox=outbox)
    ready_price.record_delivery(output, "sent", chat_id=ready_price.CHAT_ID, outbox=outbox)
    assert manifest.read_bytes() == before
    status = json.loads((manifest.parent / "telegram.json").read_text())
    assert status["ever_sent"] and len(status["attempts"]) == 2


def test_supplier_fallback_is_not_a_full_fresh_snapshot(tmp_path):
    output, source = workbook(tmp_path)
    path = ready_price.capture(output, input_path=source, supplier_fallback=True, outbox=tmp_path / "outbox")
    assert json.loads(path.read_text(encoding="utf-8"))["snapshot_kind"] == "partial"


def test_other_recipient_is_not_the_price_channel(tmp_path):
    output, source = workbook(tmp_path)
    outbox = tmp_path / "outbox"
    path = ready_price.capture(output, input_path=source, outbox=outbox)
    ready_price.record_delivery(output, "sent", chat_id="some-other-recipient", outbox=outbox)
    assert not (path.parent / "telegram.json").exists()


def test_raw_supplier_cannot_enter_outbox(tmp_path):
    _, source = workbook(tmp_path)
    with pytest.raises(ValueError):
        ready_price.capture(source, input_path=source, outbox=tmp_path / "outbox")


def test_dispatch_failure_leaves_committed_event(tmp_path, monkeypatch):
    output, source = workbook(tmp_path)
    def fail():
        raise OSError("simulated")
    monkeypatch.setattr(ready_price, "dispatch", fail)
    path = ready_price.after_transform(output, input_path=source, outbox=tmp_path / "outbox")
    assert path.is_file()


def test_telegram_timeout_is_recorded_without_token_output(monkeypatch, capsys):
    calls = []
    def fail(*args):
        raise TimeoutError("https://api.telegram.org/botSECRET/sendDocument")
    monkeypatch.setattr(telegram_send, "_send_file", fail)
    monkeypatch.setattr(ready_price, "record_delivery", lambda *a, **k: calls.append((a, k)))
    monkeypatch.setattr(ready_price, "dispatch", lambda: None)
    assert telegram_send.send_file("price.xlsx", chat_id=ready_price.CHAT_ID) is False
    assert calls[0][0][1] == "failed" and calls[0][1]["error_type"] == "TimeoutError"
    assert "SECRET" not in capsys.readouterr().out
