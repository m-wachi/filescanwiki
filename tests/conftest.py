import subprocess
import time
from pathlib import Path
import importlib

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker


COMPOSE_FILE = Path(__file__).parent / "docker-compose.postgres.yml"
DB_USER = "tracuser"
DB_PASS = "tracuser"
DB_NAME = "trac"
DB_HOST = "localhost"
DB_PORT = 5433


def wait_for_postgres(db_url, timeout=60):
    start = time.time()
    while True:
        try:
            engine = create_engine(db_url)
            with engine.connect():
                return True
        except Exception:
            if time.time() - start > timeout:
                return False
            time.sleep(0.5)


@pytest.fixture(scope="session")
def postgres_service():
    """docker-compose で Postgres を起動し、テスト用 DB URL を返す。

    起動に失敗した場合は pytest をスキップする。
    """
    # docker-compose が無い場合はスキップ
    try:
        subprocess.check_call(["docker-compose", "-v"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pytest.skip("docker-compose is required for integration tests")

    # 起動
    subprocess.check_call(["docker-compose", "-f", str(COMPOSE_FILE), "up", "-d"])

    db_url = f"postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

    if not wait_for_postgres(db_url, timeout=60):
        subprocess.check_call(["docker-compose", "-f", str(COMPOSE_FILE), "down"])
        pytest.skip("Postgres did not become ready")

    # モジュールをインポートして engine/SessionLocal を差し替え、スキーマ作成
    mod = importlib.import_module("tracticket2file01")
    mod.engine = create_engine(db_url)
    mod.SessionLocal = sessionmaker(bind=mod.engine, autoflush=False, autocommit=False)
    mod.Base.metadata.create_all(mod.engine)

    yield db_url

    # teardown
    subprocess.check_call(["docker-compose", "-f", str(COMPOSE_FILE), "down", "-v"])

