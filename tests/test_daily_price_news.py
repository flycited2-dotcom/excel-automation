import json
from pathlib import Path
import sys
from types import SimpleNamespace

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import daily_price_news as news
import telegram_send


def snapshot(folder, day, rows, fallback=False):
    path = folder / f'БытТехОпт_{day}.xlsx'
    path.write_bytes(day.encode())
    frame = pd.DataFrame([dict(source='supplier', article=article, brand='Brand',
                               name=name, group=group, price_out=price)
                          for article, name, group, price in rows])
    news.write_snapshot(frame, path, supplier_fallback=fallback)
    return path


def test_first_issue_is_neutral_then_detects_real_price_drop_and_new_item(tmp_path):
    state = tmp_path / 'state'
    old = snapshot(tmp_path, '20261012', [('1', 'Model <A>', 'Холодильники', 10000)])
    note, payload = news.prepare_note(old, 'channel', state_dir=state)
    assert 'снизились' not in note.lower() and 'появились' not in note.lower()
    news.record_success(payload, 'channel', state_dir=state)
    current = snapshot(tmp_path, '20261013', [('1', 'Model <A>', 'Холодильники', 9000),
                                           ('2', 'Generator', 'Генераторы', 15000)])
    note, payload = news.prepare_note(current, 'channel', state_dir=state)
    assert '10 000 → 9 000' in note and 'Model &lt;A&gt;' in note
    assert 'Генераторы' in note and 'поступил' not in note.lower()
    assert '12.10.2026' in note
    # Preparing a retry must not erase the preceding comparison baseline.
    assert json.loads(news.state_path('channel', state).read_text())['date'] == '2026-10-12'
    assert news.prepare_note(current, 'channel', state_dir=state)[0] == note
    news.record_success(payload, 'channel', state_dir=state)
    assert news.prepare_note(current, 'channel', state_dir=state)[0] == note


def test_no_false_claim_when_supplier_is_stale(tmp_path):
    old = snapshot(tmp_path, '20261012', [('1', 'AC', 'Кондиционеры', 100)])
    _, payload = news.prepare_note(old, 'channel', state_dir=tmp_path / 'state')
    news.record_success(payload, 'channel', state_dir=tmp_path / 'state')
    current = snapshot(tmp_path, '20261013', [('1', 'AC', 'Кондиционеры', 50)], fallback=True)
    note, _ = news.prepare_note(current, 'channel', state_dir=tmp_path / 'state')
    assert 'предыдущего сохранённого' in note and 'Снизились' not in note


def test_snapshot_must_match_the_attached_workbook(tmp_path):
    path = snapshot(tmp_path, '20261013', [('1', 'AC', 'Кондиционеры', 100)])
    path.write_bytes(b'different file')
    with pytest.raises(ValueError):
        news.prepare_note(path, 'channel')


def test_daily_focus_changes_without_inventing_changes(tmp_path):
    rows = [('1', 'AC', 'Кондиционеры', 100), ('2', 'TV', 'Телевизоры', 200)]
    first = snapshot(tmp_path, '20261012', rows)
    second = snapshot(tmp_path, '20261013', rows)
    a, _ = news.prepare_note(first, 'channel', state_dir=tmp_path / 'state')
    b, _ = news.prepare_note(second, 'channel', state_dir=tmp_path / 'state')
    assert a != b
    assert all('поступил' not in text.lower() and 'снизились' not in text.lower() for text in [a,b])


@pytest.mark.parametrize('status', [200, 429])
def test_delivery_button_and_state_only_after_success(tmp_path, monkeypatch, status):
    import requests
    path = snapshot(tmp_path, '20261013', [('1', 'AC', 'Кондиционеры', 100)])
    config = {'company': {'name': 'Company'}, 'telegram': {'bot_token': 'TEST',
              'channel_id': 'channel', 'manager_url': 'https://t.me/B2B_opt_simf'}}
    monkeypatch.setattr(telegram_send, 'load_config', lambda: config)
    monkeypatch.setattr(news, 'BASE', tmp_path)
    calls = []
    def post(url, **kwargs):
        calls.append(kwargs)
        return SimpleNamespace(status_code=status, json=lambda: {'description': 'failed'}, text='failed')
    monkeypatch.setattr(requests, 'post', post)
    assert telegram_send._send_file(str(path)) == (status == 200)
    markup = json.loads(calls[0]['data']['reply_markup'])
    assert markup['inline_keyboard'][0][0] == {'text': '💬 Написать менеджеру', 'url': 'https://t.me/B2B_opt_simf', 'style': 'success'}
    assert '13.10.2026' in calls[0]['data']['caption']
    assert news.state_path('channel').is_file() == (status == 200)
    telegram_send._send_file(str(path), caption='Supplier', chat_id='private')
    assert 'reply_markup' not in calls[-1]['data']
    assert calls[-1]['data']['caption'] == 'Supplier'
