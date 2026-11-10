from sqlalchemy import create_engine, Column, String, BigInteger, Integer
from sqlalchemy.orm import sessionmaker, declarative_base
import os, sys, datetime, configparser

# --- グローバル設定 ---


# 設定ファイルの読み込み
config = configparser.ConfigParser()
config.read('config.ini')
configDefault = config["DEFAULT"]


# PostgreSQLの接続文字列
#DB_URL = "postgresql://tracuser:tracuser@localhost:5432/trac16"
#DB_URL = "postgresql://tracuser:tracuser@localhost:5432/trac"
DB_URL = configDefault["tracdb_url"]


# 📂 ファイル保存先フォルダ
#OUTPUT_DIR = "c:\\tmp\\wiki_exports" 
#OUTPUT_DIR = "/tmp/wiki_exports"
OUTPUT_DIR = configDefault["output_dir"]


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
    file_name_only = f"{target_name}.tracwiki.txt"
    output_filepath = os.path.join(OUTPUT_DIR, "wiki", file_name_only)
    
    session = Session()
    
    print(f"\n--- 処理対象: {target_name} ---")


    # target_nameに対応する最新のWikiレコードを取得
    rec_wiki = get_latest_wiki(session, target_name)

    if rec_wiki:
        # 内容は変換せずそのまま書き出す
        file_content : str = rec_wiki.text

        #print(f"file_content={file_content}")
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
            # 出力先ディレクトリが無ければ作成（ページ名に"/"が含まれる場合も考慮）
            os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
            # newline='' で改行コードを変換せずそのまま書き出す
            with open(output_filepath, 'w', encoding='utf-8', newline='') as f:
                f.write(file_content)
            
            print(f"ファイル {output_filepath} を正常に更新しました。")
            
        else:
            db_time_display = datetime.datetime.fromtimestamp(db_timestamp) if db_timestamp else "N/A"
            print("更新がないため、書き出しをスキップします。")

    else:
        print(f"⚠️ '{target_name}' のレコードは見つかりませんでした。")

    session.close()


def read_wiki_page_names(list_path: str):
    """
    引数で指定されたファイルからWikiPageNameを読み込む（改行区切り、複数可）。
    空行は無視し、前後の空白は除去する。
    """
    with open(list_path, 'r', encoding='utf-8-sig') as f:
        return [line.strip() for line in f if line.strip()]


if __name__ == "__main__":

    if len(sys.argv) != 2:
        print(f"使い方: python {os.path.basename(sys.argv[0])} <WikiPageName一覧ファイル>")
        sys.exit(1)

    # 保存先フォルダの確認と作成
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"保存先フォルダ '{OUTPUT_DIR}' を作成しました。")

    for page_name in read_wiki_page_names(sys.argv[1]):
        save_wiki_to_file(page_name)
