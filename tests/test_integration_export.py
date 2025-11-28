import pytest
from pathlib import Path
import importlib


def test_export_with_postgres(postgres_service, tmp_path: Path):
    mod = importlib.import_module("tracticket2file01")

    sess = mod.SessionLocal()
    # 簡単なレコードを挿入
    t = mod.Ticket(id=1, summary="Integration test", status="new")
    sess.add(t)
    sess.add(mod.TicketChange(ticket=1, time=1, author="x", field="status", oldvalue="new", newvalue="open"))
    sess.commit()

    cfg = tmp_path / "config.ini"
    cfg.write_text(f"[DEFAULT]\ntracdb_url = {postgres_service}\noutput_dir = {tmp_path}\n")

    out = mod.export_ticket_to_file(1, config_path=cfg)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "Integration test" in content

