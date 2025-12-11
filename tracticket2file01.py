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
import configparser, datetime, os

from sqlalchemy import (
    Column,
    Integer,
    BigInteger,
    Text,
    create_engine,
    PrimaryKeyConstraint,
)
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.sql import func

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

def get_ticket_max_id():
    """
    チケット番号の最大値を取得する
    """
    id_max = 0
    with get_session() as sess:
        rec = sess.query(func.max(Ticket.id).label('id_max')).one_or_none()
        id_max = rec.id_max
    return id_max

def is_file_older(filename: str, ticket_update_timestamp: float) -> bool:
    """
    ファイルがチケットの更新日時と比較して古い状態にある（上書きが必要）かを判断する。

    ファイルが存在しない場合、またはチケットの更新日時がファイルの更新日時と等しいか新しい場合に True を返す。
    
    Args:
        filename (str): 出力先ファイルのパス（文字列）。
        ticket_update_timestamp (float): チケットの最終更新日時のUNIXタイムスタンプ（秒）。

    Returns:
        bool: ファイルがチケットより古い状態にあるか（上書きが必要）であれば True、そうでなければ False。
    """
    filepath = Path(filename) # 文字列から Path オブジェクトに変換
    
    # 1. ファイルが存在するか？
    if not filepath.exists():
        # ファイルが存在しない場合は、古い（更新が必要）と見なす (True)
        return True
    
    # 2. ファイルが存在する場合、ファイルの更新日時を取得
    try:
        # st_mtime: 最終修正時刻 (UNIXタイムスタンプ/秒)
        file_mtime_timestamp = filepath.stat().st_mtime
    except OSError:
        # ファイルの stat 取得に失敗した場合、安全を見て古い（更新が必要）と見なす (True)
        return True
    
    # 3. チケットの更新日時が、ファイルの更新日時と等しいか、より新しいか？
    # チケット更新日時 (ticket_update_timestamp) >= ファイル更新日時 (file_mtime_timestamp) の場合は、
    # ファイルが古いか同等と見なし、更新が必要 (True)
    if ticket_update_timestamp >= file_mtime_timestamp:
        return True
    else:
        # チケット更新日時 < ファイル更新日時 (ファイルの方が新しい) の場合は False
        return False


def export_ticket_to_file(ticket_no) -> str:
    """指定したチケット番号の `ticket` と `ticket_change` をテキストファイルに書き出す。

    ファイル名は `ticket_<No>.txt`、保存先は `config.ini` の `output_dir`。
    存在しないチケットでも、該当する変更履歴があればそれを書き出す。
    出力先ファイルが存在し、その更新日時がチケットの更新日時より新しい場合は、ファイルを上書きしない。

    戻り値は書き出したファイルのパス (Path)。ファイルを書き出さなかった場合はNone
    """

    filename = f"{OUTPUT_DIR}/ticket_{ticket_no}.txt"

    # DB からデータ取得
    with get_session() as sess:
        try:
            ticket_obj = sess.query(Ticket).filter(Ticket.id == int(ticket_no)).one_or_none()
        except Exception:
            ticket_obj = None
        
        # チケットの更新日時を取得（ファイル比較のためにUNIXタイムスタンプとして扱う）
        ticket_update_timestamp = 0
        if ticket_obj and ticket_obj.changetime is not None:
            # Tracのタイムスタンプ（マイクロ秒）を datetime に変換
            ticket_update_dt = tractime_to_datetime(ticket_obj.changetime)
            # UNIXタイムスタンプ（秒）に変換
            ticket_update_timestamp = ticket_update_dt.timestamp() 
        
        changes = (
            sess.query(TicketChange)
            .filter(TicketChange.ticket == int(ticket_no))
            .order_by(TicketChange.oldvalue)
            .all()
        )

    # --- 上書き条件のチェック ---
    # is_file_older() が True (ファイルが古い/不在) の場合のみ書き込みを行う
    if not is_file_older(filename, ticket_update_timestamp):
        # ファイルの更新日時がチケットより新しいので、上書きせずに終了
        # print(f"File {filename} is newer. Skipping write.")
        return None
            
    # --- ファイルへ書き込み（上書きが必要な場合） ---
    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"Ticket: {ticket_no}\n")
        f.write("=" * 60 + "\n")
        
        # ticket_obj がある場合、内容を書き出す
        if ticket_obj:
            f.write(f"{ticket_obj.summary}\n")
            f.write("-" * 60 + "\n")
            
            # 登録日時と更新日時
            vtmp = tractime_to_datetime(ticket_obj.time)
            f.write(f"  登録日時: {vtmp:%Y/%m/%d %H:%M}, ")
            vtmp = tractime_to_datetime(ticket_obj.changetime)
            f.write(f"更新日時: {vtmp:%Y/%m/%d %H:%M} \n")
            
            f.write(f"  Milestone: {ticket_obj.milestone}\n")
            f.write("\n")
            f.write(f"詳細:\n{ticket_obj.description}\n")
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

        # 変更履歴（コメント）を書き出す
        if changes:
            for ch in changes:
                if ch.field != "comment":
                    continue
                vtmp = tractime_to_datetime(ch.time)
                f.write(f"{ch.oldvalue} -  {vtmp:%Y/%m/%d %H:%M}\n\n")
                f.write(f"{ch.newvalue}\n")
                f.write("-" * 60 + "\n")
        else:
            f.write("No changes found.\n")

    return filename


if __name__ == "__main__":
    # 簡易チェック: DB URLを表示
    print("Database URL:", DATABASE_URL)
    print("Defined tables:", Base.metadata.tables.keys())

    # チケット番号の最大値を取得
    id_max = get_ticket_max_id()
    print(f"ticket max id={id_max}\n")

    # チケット番号最大値まで順にループを回してファイルに出力する
    # ファイルの更新日時と比較してチケットの更新日時が古い
    # ファイルは書き出さない
    for n in range(1, id_max+1):
        out_file = export_ticket_to_file(n)
        if out_file:
            print(f"Exported ticket {n} -> {out_file}")

