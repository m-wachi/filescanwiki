"""
Trac の `ticket` と `ticket_change` テーブル用の SQLAlchemy モデル。

このモジュールは `config.ini`（DEFAULT.tracdb_url）からデータベース接続
URL を `configparser` で読み取り、SQLAlchemy のエンジンを作成し、
宣言的ベース `Base`、モデルクラス、および `get_session()` ヘルパーを提供します。

使用例:
  from tracticket2file01 import get_session, Ticket, TicketChange
  with get_session() as sess:
      tickets = sess.query(Ticket).limit(10).all()
"""
from pathlib import Path
import configparser

from sqlalchemy import (
    Column,
    Integer,
    BigInteger,
    Text,
    create_engine,
    PrimaryKeyConstraint,
)
from sqlalchemy.orm import declarative_base, sessionmaker


BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.ini"


def load_db_url(config_path: Path = CONFIG_PATH) -> str:
    """`config.ini` の DEFAULT セクションから `tracdb_url` を読み取る。

    キーが存在しない場合は `KeyError` を送出します。
    """
    cfg = configparser.ConfigParser()
    read_files = cfg.read(config_path)
    if not read_files:
        raise FileNotFoundError(f"Config file not found: {config_path}")
    try:
        return cfg["DEFAULT"]["tracdb_url"]
    except KeyError:
        raise KeyError("'tracdb_url' missing in DEFAULT section of config.ini")


# SQLAlchemy の設定
DATABASE_URL = load_db_url()
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()


class Ticket(Base):
    """`ticket` テーブルのモデル。

    `doc/database.txt` に基づくカラム割り当て:
      - id        serial4 -> Integer (主キー)
      - type      text
      - time      int8 -> BigInteger
      - changetime int8 -> BigInteger
      - component text
      - severity  text
      - priority  text
      - owner     text
      - reporter  text
      - cc        text
      - version   text
      - milestone text
      - status    text
      - resolution text
      - summary   text
      - description text
      - keywords  text
    """

    __tablename__ = "ticket"

    id = Column(Integer, primary_key=True)
    type = Column(Text)
    time = Column(BigInteger)
    changetime = Column(BigInteger)
    component = Column(Text)
    severity = Column(Text)
    priority = Column(Text)
    owner = Column(Text)
    reporter = Column(Text)
    cc = Column(Text)
    version = Column(Text)
    milestone = Column(Text)
    status = Column(Text)
    resolution = Column(Text)
    summary = Column(Text)
    description = Column(Text)
    keywords = Column(Text)

    def __repr__(self) -> str:
        return (
            f"<Ticket id={self.id!r} summary={self.summary!r} "
            f"status={self.status!r} owner={self.owner!r}>"
        )

    def to_dict(self) -> dict:
        """この Ticket のカラム名をキー、値をバリューとする辞書を返す。"""
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}


class TicketChange(Base):
    """`ticket_change` テーブルのモデル。

    元のスキーマには単一の主キー列が記載されていなかったため、
    チケットの変更イベントの一意性を反映して `(ticket, time, field)` の
    複合主キーを使用する。
    """

    __tablename__ = "ticket_change"
    __table_args__ = (PrimaryKeyConstraint("ticket", "time", "field"),)

    ticket = Column(Integer)
    time = Column(BigInteger)
    author = Column(Text)
    field = Column(Text)
    oldvalue = Column(Text)
    newvalue = Column(Text)

    def __repr__(self) -> str:
        return (
            f"<TicketChange ticket={self.ticket!r} time={self.time!r} "
            f"field={self.field!r} author={self.author!r}>"
        )

    def to_dict(self) -> dict:
        """この TicketChange のカラム名をキー、値をバリューとする辞書を返す。"""
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}


def get_session():
    """コンテキスト管理で使用できる SQLAlchemy セッションを返す。

    使用例:
      with get_session() as sess:
          ...
    """
    return SessionLocal()


def load_output_dir(config_path: Path = CONFIG_PATH) -> Path:
    """`config.ini` の DEFAULT セクションから `output_dir` を読み取り、Path を返す。

    存在しない場合はディレクトリを作成する。
    """
    cfg = configparser.ConfigParser()
    read_files = cfg.read(config_path)
    if not read_files:
        raise FileNotFoundError(f"Config file not found: {config_path}")
    try:
        out = cfg["DEFAULT"]["output_dir"]
    except KeyError:
        raise KeyError("'output_dir' missing in DEFAULT section of config.ini")
    out_path = Path(out)
    out_path.mkdir(parents=True, exist_ok=True)
    return out_path


def export_ticket_to_file(ticket_no, config_path: Path = CONFIG_PATH) -> Path:
    """指定したチケット番号の `ticket` と `ticket_change` をテキストファイルに書き出す。

    ファイル名は `ticket_<No>.txt`、保存先は `config.ini` の `output_dir`。
    存在しないチケットでも、該当する変更履歴があればそれを書き出す。

    戻り値は書き出したファイルのパス (Path)。
    """
    out_dir = load_output_dir(config_path)
    filename = out_dir / f"ticket_{ticket_no}.txt"

    # DB からデータ取得
    with get_session() as sess:
        try:
            ticket_obj = sess.query(Ticket).filter(Ticket.id == int(ticket_no)).one_or_none()
        except Exception:
            ticket_obj = None
        changes = (
            sess.query(TicketChange)
            .filter(TicketChange.ticket == int(ticket_no))
            .order_by(TicketChange.time)
            .all()
        )

    # ファイルへ書き込み
    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"Ticket: {ticket_no}\n")
        f.write("=" * 60 + "\n")
        if ticket_obj:
            for k, v in ticket_obj.to_dict().items():
                f.write(f"{k}: {v}\n")
        else:
            f.write("Ticket not found.\n")

        f.write("\nChanges:\n")
        f.write("=" * 60 + "\n")
        if changes:
            for ch in changes:
                # 単純フォーマット: time, author, field, old -> new
                f.write(
                    f"time={ch.time} author={ch.author} field={ch.field} \n"
                )
                f.write(f"  old: {ch.oldvalue}\n")
                f.write(f"  new: {ch.newvalue}\n")
                f.write("-" * 40 + "\n")
        else:
            f.write("No changes found.\n")

    return filename


if __name__ == "__main__":
    # 簡易チェック: DB URL と定義されているテーブルを表示し、
    # チケット番号 1 と 2 をファイルに書き出す
    print("Database URL:", DATABASE_URL)
    print("Defined tables:", Base.metadata.tables.keys())

    for n in (1, 2):
        try:
            out_path = export_ticket_to_file(n)
            print(f"Exported ticket {n} -> {out_path}")
        except Exception as exc:
            print(f"Failed to export ticket {n}: {exc}")
