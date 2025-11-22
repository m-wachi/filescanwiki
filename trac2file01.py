from sqlalchemy import create_engine, Column, String, BigInteger, Integer
from sqlalchemy.orm import sessionmaker, declarative_base
import os
import datetime
import traceback
import sys

# --- グローバル設定 ---
# PostgreSQLの接続文字列
#DB_URL = "postgresql://tracuser:tracuser@localhost:5432/trac16"
DB_URL = "postgresql://tracuser:tracuser@localhost:5432/trac"

# 📂 ファイル保存先フォルダ
#OUTPUT_DIR = "c:\\tmp\\wiki_exports" 
OUTPUT_DIR = "/tmp/wiki_exports"

# --- データベースのセットアップ ---
# client_encoding='utf8' を指定してしないと日本語データでエラーになる
engine = create_engine(DB_URL, client_encoding='utf8')
Base = declarative_base()

# テーブルに対応するWikiクラスを定義
class Wiki(Base):
    __tablename__ = 'wiki'
    name = Column(String, primary_key=True) 
    version = Column(Integer, primary_key=True)
    text = Column(String)                   
    time = Column(BigInteger)               

    def __repr__(self):
        return f"<Wiki(name='{self.name}', time={self.time})>"

# セッションファクトリの作成
Session = sessionmaker(bind=engine)


def get_latest_wiki(session, target_name: str):
    """指定されたnameの最新バージョンのWikiレコードを取得する関数。

    Args:
        session: SQLAlchemyのセッションオブジェクト。
        target_name (str): 検索対象となるwikiテーブルのnameカラムの値。
    Returns:
        最新バージョンのWikiレコードオブジェクト。
        もし該当レコードが存在しない場合はNoneを返す。
    """
    result = session.query(Wiki).filter_by(name=target_name).order_by(Wiki.version.desc())
    if result.count() == 0:
        return None
    return result.first()

def tractime_to_timestamp(tractime):
    """Tracのtimeカラムの値をUNIXタイムスタンプに変換する関数。

    Args:
        tractime (int): Tracのtimeカラムの値（マイクロ秒単位）。

    Returns:
        int: UNIXタイムスタンプ（秒単位）。
    """
    return tractime / 1000000  # マイクロ秒を秒に変換

def get_file_modification_time(filepath: str) -> int:
    """指定されたファイルの最終更新日時を取得する関数。

    Args:
        filepath (str): ファイルのパス。

    Returns:
        int: 最終更新日時。（os.stat(filepath).st_mtimeを利用）
    """
    return os.stat(filepath).st_mtime

# --- メイン処理 ---
def save_wiki_to_file(target_name: str):
    """
    SQLAlchemy ORMを使用して、timeカラムとファイルのタイムスタンプを比較し、
    新しい場合にのみデータを取得・保存する関数。

    Args:
        target_name (str): 検索対象となるwikiテーブルのnameカラムの値。
    """
   
    # 2. フルパスのファイル名を生成
    file_name_only = f"{target_name}.txt"
    output_filepath = os.path.join(OUTPUT_DIR, file_name_only)
    
    session = Session()
    
    print(f"\n--- 処理対象: {target_name} ---")


    # target_nameに対応する最新のWikiレコードを取得
    rec_wiki = get_latest_wiki(session, target_name)

    if rec_wiki:
        file_content = rec_wiki.text
        db_timestamp = tractime_to_timestamp(rec_wiki.time)
        print(f"DBのtime: {datetime.datetime.fromtimestamp(db_timestamp)}")


        # 既存ファイルの最終更新UNIXタイムスタンプを取得
        if os.path.exists(output_filepath):
            file_timestamp = get_file_modification_time(output_filepath)
            print(f"既存ファイルのtime: {datetime.datetime.fromtimestamp(file_timestamp)}")
        else:
            file_timestamp = 0
            print(f"対象ファイル '{file_name_only}' は存在しません。新規書き込みを行います。")


        # 5. タイムスタンプの比較
        if db_timestamp > file_timestamp:
            print("データベースのデータが新しいため、ファイルを上書きします。")

            # 6. ファイルへの書き出し（上書き）
            with open(output_filepath, 'w', encoding='utf-8') as f:
                f.write(file_content)
            
            print(f"ファイル {output_filepath} を正常に更新しました。")
            
        else:
            db_time_display = datetime.datetime.fromtimestamp(db_timestamp) if db_timestamp else "N/A"
            print("更新がないため、書き出しをスキップします。")

    else:
        print(f"⚠️ '{target_name}' のレコードは見つかりませんでした。")



if __name__ == "__main__":

    # 保存先フォルダの確認と作成
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"保存先フォルダ '{OUTPUT_DIR}' を作成しました。")


    # 呼び出し例 1
    FIRST_TARGET = '__MapPage00002754'
    save_wiki_to_file(FIRST_TARGET)
    
    print("\n" + "="*40)

    # 呼び出し例 2
    SECOND_TARGET = 'AnotherWikiPage'
    save_wiki_to_file(SECOND_TARGET)
   
