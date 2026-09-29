"""Атомарная доставка принятых выпусков на серверный inbox через OpenSSH."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
import subprocess
import tarfile
import tempfile


def push_release(release: Path, *, host="sprintbox-itp", remote_root="/opt/avito-ready-price"):
    meta = json.loads(release.read_text(encoding="utf-8"))
    digest = meta["source_sha256"]
    if hashlib.sha256((release.parent / "price.xlsx").read_bytes()).hexdigest() != digest:
        raise ValueError("release checksum mismatch before upload")
    supplier_file = meta.get("supplier_file")
    if supplier_file:
        if supplier_file not in {"supplier.xls", "supplier.xlsx"}:
            raise ValueError("invalid supplier file")
        if hashlib.sha256((release.parent / supplier_file).read_bytes()).hexdigest() != meta["provenance"]["supplier"]["sha256"]:
            raise ValueError("supplier checksum mismatch before upload")
    if not __import__("re").fullmatch(r"[0-9a-f]{64}", digest):
        raise ValueError("invalid release digest")
    with tempfile.TemporaryDirectory() as temp:
        bundle = Path(temp) / f"{digest}.tgz"
        with tarfile.open(bundle, "w:gz") as archive:
            for name in ("release.json", "price.xlsx", supplier_file, "telegram.json"):
                if not name:
                    continue
                path = release.parent / name
                if path.is_file():
                    archive.add(path, arcname=name)
        partial = f"{remote_root}/incoming/{digest}.tgz.partial"
        final = f"{remote_root}/inbox/{digest}.tgz"
        subprocess.run(["ssh", "-o", "BatchMode=yes", host,
                        f"umask 077; mkdir -p {remote_root}/incoming {remote_root}/inbox {remote_root}/archive"],
                       check=True, timeout=30)
        subprocess.run(["scp", "-q", str(bundle), f"{host}:{partial}"], check=True, timeout=120)
        subprocess.run(["ssh", "-o", "BatchMode=yes", host,
                        f"test -f {final} || mv {partial} {final}; rm -f {partial}"], check=True, timeout=30)
    return digest
