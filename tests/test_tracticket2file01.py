import importlib
from pathlib import Path

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

import tracticket2file01 as mod


def setup_inmemory_db():
    """モジュールの engine/SessionLocal を in-memory SQLite に差し替え、テーブルを作成する。"""
    engine = create_engine("sqlite:///:memory:")
    mod.engine = engine
    mod.SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
    mod.Base.metadata.create_all(engine)
    return engine


def test_export_ticket_to_file_writes_file(tmp_path: Path):
    # in-memory DB に差し替え
    setup_inmemory_db()

    # サンプルデータ挿入
    sess = mod.SessionLocal()
    t = mod.Ticket(
        id=1,
        type="defect",
        time=1600000000,
        changetime=1600000100,
        component="core",
        severity="major",
        priority="high",
        owner="alice",
        reporter="bob",
        cc="",
        version="1.0",
        milestone="",
        status="new",
        resolution="",
        summary="Test ticket",
        description="This is a test.",
        keywords="test",
    )
    sess.add(t)
    ch = mod.TicketChange(ticket=1, time=1600000200, author="carol", field="status", oldvalue="new", newvalue="assigned")
    sess.add(ch)
    sess.commit()

    # 一時的な config.ini を用意して出力先を tmp_path に設定
    cfg_path = tmp_path / "config.ini"
    cfg_path.write_text(f"[DEFAULT]\ntracdb_url = sqlite:///:memory:\noutput_dir = {tmp_path}\n")

    out_file = mod.export_ticket_to_file(1, config_path=cfg_path)
    assert out_file.exists()

    content = out_file.read_text(encoding="utf-8")
    assert "summary: Test ticket" in content
    assert "author=carol" in content or "carol" in content
    assert "old: new" in content
    assert "new: assigned" in content


def test_export_nonexistent_ticket_writes_changes(tmp_path: Path):
    # in-memory DB に差し替え
    setup_inmemory_db()

    # チケット本体がないが変更のみ存在するケース
    sess = mod.SessionLocal()
    ch = mod.TicketChange(ticket=2, time=1600000300, author="dave", field="priority", oldvalue="low", newvalue="high")
    sess.add(ch)
    sess.commit()

    cfg_path = tmp_path / "config.ini"
    cfg_path.write_text(f"[DEFAULT]\ntracdb_url = sqlite:///:memory:\noutput_dir = {tmp_path}\n")

    out_file = mod.export_ticket_to_file(2, config_path=cfg_path)
    content = out_file.read_text(encoding="utf-8")
    # ticket not found メッセージが含まれ、changes は書き出される
    assert "Ticket not found" in content
    assert "author=dave" in content or "dave" in content

