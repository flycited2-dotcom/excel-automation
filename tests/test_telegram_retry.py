import sys
from pathlib import Path

import pytest
import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import telegram_send


@pytest.fixture
def calls(monkeypatch):
    monkeypatch.setattr(telegram_send, 'CONNECT_PAUSE', 0)
    log = []
    monkeypatch.setattr(telegram_send, '_fake_errors', [], raising=False)

    def fake(filepath, caption, chat_id):
        log.append(1)
        if telegram_send._fake_errors:
            raise telegram_send._fake_errors.pop(0)
        return True
    monkeypatch.setattr(telegram_send, '_send_file', fake)
    return log


def test_connect_timeout_is_retried(calls):
    """06.11.2026: одна ConnectTimeout с VPS сорвала рассылку в канал."""
    telegram_send._fake_errors[:] = [requests.exceptions.ConnectTimeout('x')] * 2
    assert telegram_send._send_with_retry('f.xlsx', None, None) is True
    assert len(calls) == 3


def test_read_timeout_is_not_retried_to_avoid_duplicates(calls):
    telegram_send._fake_errors[:] = [requests.exceptions.ReadTimeout('x')]
    with pytest.raises(requests.exceptions.ReadTimeout):
        telegram_send._send_with_retry('f.xlsx', None, None)
    assert len(calls) == 1


def test_aborted_upload_is_not_retried(calls):
    telegram_send._fake_errors[:] = [requests.exceptions.ConnectionError('Connection aborted')]
    with pytest.raises(requests.exceptions.ConnectionError):
        telegram_send._send_with_retry('f.xlsx', None, None)
    assert len(calls) == 1


def test_gives_up_after_all_attempts(calls):
    telegram_send._fake_errors[:] = [requests.exceptions.ConnectTimeout('x')] * 10
    with pytest.raises(requests.exceptions.ConnectTimeout):
        telegram_send._send_with_retry('f.xlsx', None, None)
    assert len(calls) == telegram_send.CONNECT_ATTEMPTS
