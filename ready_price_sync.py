"""Повторяемый обработчик локального outbox. Сеть и публикация пока не включены."""
import hashlib
import json
from pathlib import Path
import sys

from ready_price import BASE_DIR, locked, utcnow, write_json


def sync(config_path=None):
    config_path = Path(config_path) if config_path else BASE_DIR / "ready-price-bridge.json"
    cfg = json.loads(config_path.read_text(encoding="utf-8"))
    bridge = Path(cfg["bridge_root"])
    sys.path.insert(0, str(bridge / "src"))
    from avito_bridge.ready_price.store import export_catalog, import_release
    outbox = Path(cfg["outbox"])
    database, categories = Path(cfg["database"]), Path(cfg["categories"])
    results = []
    with locked(outbox / ".sync"):
        # Сначала старый выпуск, затем новый: стабильное восстановление после простоя.
        manifests = []
        for path in outbox.glob("*/release.json"):
            try:
                metadata = json.loads(path.read_text(encoding="utf-8"))
                manifests.append((metadata.get("generated_at", ""), path))
            except (ValueError, OSError):
                results.append({"event": path.parent.name, "status": "error", "reason": "invalid_release_json"})
        for _, manifest in sorted(manifests):
            receipt = manifest.parent / "local-import.json"
            telegram = manifest.parent / "telegram.json"
            fingerprint = hashlib.sha256(manifest.read_bytes() + categories.read_bytes()
                                          + (telegram.read_bytes() if telegram.exists() else b"")).hexdigest()
            try:
                if receipt.exists():
                    previous = json.loads(receipt.read_text(encoding="utf-8"))
                    if previous.get("input_fingerprint") == fingerprint and previous.get("status") in {"accepted", "quarantined"}:
                        continue
                result = import_release(manifest, database, categories)
                write_json(receipt, {**result, "input_fingerprint": fingerprint, "checked_at": utcnow()})
                results.append(result)
            except Exception as exc:
                # Не логируем потенциально чувствительные URL/тексты исключений.
                result = {"event": manifest.parent.name, "status": "error", "reason": type(exc).__name__, "checked_at": utcnow()}
                write_json(receipt, result)
                results.append(result)
        if cfg.get("remote_push", {}).get("enabled"):
            from ready_price_remote import push_release
            remote = cfg["remote_push"]
            for _, manifest in sorted(manifests):
                receipt = manifest.parent / "local-import.json"
                pushed = manifest.parent / "remote-push.json"
                if pushed.exists() or not receipt.exists():
                    continue
                local = json.loads(receipt.read_text(encoding="utf-8"))
                if local.get("status") != "accepted":
                    continue
                try:
                    digest = push_release(manifest, host=remote.get("host", "sprintbox-itp"),
                                          remote_root=remote.get("root", "/opt/avito-ready-price"))
                    write_json(pushed, {"sha256": digest, "status": "uploaded", "at": utcnow()})
                except Exception as exc:
                    results.append({"stage": "remote_push", "status": "error", "reason": type(exc).__name__})
        if database.exists():
            write_json(Path(cfg["catalog_export"]), export_catalog(database))
            if cfg.get("identity_references"):
                try:
                    from avito_bridge.ready_price.identity import reconcile
                    import yaml
                    refs = cfg["identity_references"]
                    owned = yaml.safe_load(Path(refs["itp_mapping"]).read_text(encoding="utf-8"))["sku_to_ad_id"].values()
                    report = reconcile(database, Path(refs["cached_feed"]), Path(refs["cached_manual_stops"]), owned)
                    write_json(Path(refs["report"]), report)
                except Exception as exc:
                    results.append({"stage": "identity", "status": "error", "reason": type(exc).__name__})
        write_json(outbox / "sync-status.json", {"checked_at": utcnow(), "results": results,
                   "remote_transfer_enabled": bool(cfg.get("remote_push", {}).get("enabled")),
                   "publication_enabled": "server_controlled"})
    print(json.dumps({"processed": len(results), "statuses": [r["status"] for r in results]}, ensure_ascii=True))
    return 1 if any(r["status"] == "error" for r in results) else 0


if __name__ == "__main__":
    raise SystemExit(sync())
