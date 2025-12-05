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
import configparser, datetime

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

cfg = configparser.ConfigParser()
cfg.read(BASE_DIR / "config.ini")
cfgDefault = cfg["DEFAULT"]

OUTPUT_DIR = cfgDefault["output_dir"]
DATABASE_URL = cfgDefault["tracdb_url"]


# SQLAlchemy の設定
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()


class Ticket(Base):
    """`ticket` テーブルのモデル。
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


def tractime_to_timestamp(tractime):
    """Tracのtimeカラムの値をUNIXタイムスタンプに変換する関数。

    Args:
        tractime (int): Tracのtimeカラムの値（マイクロ秒単位）。

    Returns:
        int: UNIXタイムスタンプ（秒単位）。
    """
    return tractime / 1000000  # マイクロ秒を秒に変換

def tractime_to_datetime(tractime):
    tmstmp = tractime_to_timestamp(tractime)
    return datetime.datetime.fromtimestamp(tmstmp)

def export_ticket_to_file(ticket_no) -> Path:
    """指定したチケット番号の `ticket` と `ticket_change` をテキストファイルに書き出す。

    ファイル名は `ticket_<No>.txt`、保存先は `config.ini` の `output_dir`。
    存在しないチケットでも、該当する変更履歴があればそれを書き出す。

    戻り値は書き出したファイルのパス (Path)。
    """
    filename = f"{OUTPUT_DIR}/ticket_{ticket_no}.txt"

    # DB からデータ取得
    with get_session() as sess:
        try:
            ticket_obj = sess.query(Ticket).filter(Ticket.id == int(ticket_no)).one_or_none()
        except Exception:
            ticket_obj = None
        changes = (
            sess.query(TicketChange)
            .filter(TicketChange.ticket == int(ticket_no))
            .order_by(TicketChange.oldvalue)
            .all()
        )

    # ファイルへ書き込み
    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"Ticket: {ticket_no}\n")
        f.write("=" * 60 + "\n")
        if ticket_obj:
            f.write(f"{ticket_obj.summary}\n")
            f.write("-" * 60 + "\n")
            vtmp = tractime_to_datetime(ticket_obj.time)
            f.write(f"  登録日時: {vtmp:%Y/%m/%d %H:%M}, ")
            vtmp = tractime_to_datetime(ticket_obj.changetime)
            f.write(f"更新日時: {vtmp:%Y/%m/%d %H:%M} \n")
            f.write(f"  Milestone: {ticket_obj.milestone}\n")
            f.write("\n")
            f.write(f"詳細:\n{ticket_obj.description}\n")
            #f.write("-" * 60 + "\n")
            #for k, v in ticket_obj.to_dict().items():
            #    f.write(f"{k}: {v}\n")
        else:
            f.write("Ticket not found.\n")

        f.write("=" * 60 + "\n")
        f.write("\nComments:\n")
        f.write("-" * 60 + "\n")
        #
        # ticket_changeテーブルに入っているコメントについて
        # field = "comment"のデータがコメント
        #   oldvalue: コメント番号
        #   newvalue: コメントの内容
        # field = "_comment?"はコメントの過去データ
        #   time: 対応する最新のコメント(field=comment)のtimeと同じ値になる
        #   oldvalue: （古い）コメントの内容
        #   newvalue: （古い）コメントの更新日時
        #
        if changes:
            for ch in changes:
                if ch.field != "comment":
                    continue
                vtmp = tractime_to_datetime(ch.time)
                f.write(f"{ch.oldvalue} -  {vtmp:%Y/%m/%d %H:%M}\n\n")
                f.write(f"{ch.newvalue}\n")
                # 単純フォーマット: time, author, field, old -> new
                # f.write(
                #     f"time={ch.time} author={ch.author} field={ch.field} \n"
                # )
                # f.write(f"  old: {ch.oldvalue}\n")
                # f.write(f"  new: {ch.newvalue}\n")
                f.write("-" * 60 + "\n")
        else:
            f.write("No changes found.\n")

    return filename


if __name__ == "__main__":
    # 簡易チェック: DB URL と定義されているテーブルを表示し、
    # チケット番号 1 と 2 をファイルに書き出す
    print("Database URL:", DATABASE_URL)
    print("Defined tables:", Base.metadata.tables.keys())

    for n in (1, 2, 1297):
        out_path = export_ticket_to_file(n)
        print(f"Exported ticket {n} -> {out_path}")
