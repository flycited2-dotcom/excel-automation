import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ready_price_remote
import ready_price_sync


def release(outbox, name, *, kind="full", generated_at="2026-11-05T06:22:00+00:00"):
    event = outbox / name
    event.mkdir(parents=True)
    (event / "release.json").write_text(json.dumps(
        {"snapshot_kind": kind, "generated_at": generated_at}), encoding="utf-8")
    return event


def config(tmp_path, outbox):
    path = tmp_path / "bridge.json"
    path.write_text(json.dumps({"mode": "push_only", "outbox": str(outbox),
                                "remote_push": {"enabled": True, "host": "sprintbox-avito",
                                                "transport": "stream"}}), encoding="utf-8")
    return path


def test_push_only_sends_full_releases_once_in_order(tmp_path, monkeypatch):
    outbox = tmp_path / "outbox"
    release(outbox, "b" * 64, generated_at="2026-11-05T07:00:00+00:00")
    release(outbox, "a" * 64, generated_at="2026-11-05T06:00:00+00:00")
    release(outbox, "c" * 64, kind="partial")  # fallback-снимок не отправляется
    sent = []
    monkeypatch.setattr(ready_price_remote, "push_release",
                        lambda manifest, **kw: sent.append((manifest.parent.name, kw)) or manifest.parent.name)
    assert ready_price_sync.sync(config(tmp_path, outbox)) == 0
    assert [name for name, _ in sent] == ["a" * 64, "b" * 64]
    assert sent[0][1]["host"] == "sprintbox-avito" and sent[0][1]["transport"] == "stream"
    assert (outbox / ("a" * 64) / "remote-push.json").exists()
    assert not (outbox / ("c" * 64) / "remote-push.json").exists()
    sent.clear()
    assert ready_price_sync.sync(config(tmp_path, outbox)) == 0
    assert sent == []  # повтор ничего не отправляет


def test_push_only_error_keeps_release_for_retry(tmp_path, monkeypatch):
    outbox = tmp_path / "outbox"
    release(outbox, "a" * 64)

    def fail(manifest, **kw):
        raise OSError("network down")
    monkeypatch.setattr(ready_price_remote, "push_release", fail)
    assert ready_price_sync.sync(config(tmp_path, outbox)) == 1
    assert not (outbox / ("a" * 64) / "remote-push.json").exists()
    status = json.loads((outbox / "sync-status.json").read_text(encoding="utf-8"))
    assert status["results"][0]["reason"] == "OSError"
