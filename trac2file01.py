from sqlalchemy import create_engine, Column, String, BigInteger
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import os
import datetime

# --- グローバル設定 ---
# PostgreSQLの接続文字列
DB_URL = "postgresql://tracuser:tracuser@localhost:5432/trac16"

# 📂 ファイル保存先フォルダ
OUTPUT_DIR = "c:\\tmp\\wiki_exports" 

# --- データベースのセットアップ ---
engine = create_engine(DB_URL)
Base = declarative_base()

# テーブルに対応するWikiクラスを定義
class Wiki(Base):
    __tablename__ = 'wiki'
    name = Column(String, primary_key=True) 
    version = Column(int)
    text = Column(String)                   
    time = Column(BigInteger)               

    def __repr__(self):
        return f"<Wiki(name='{self.name}', time={self.time})>"

# セッションファクトリの作成
Session = sessionmaker(bind=engine)

# --- メイン処理 ---
def fetch_and_save_if_newer_orm(target_name: str):
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
    file_mtime = 0
    
    print(f"\n--- 処理対象: {target_name} ---")

    # 3. 既存ファイルの最終更新UNIXタイムスタンプを取得
    if os.path.exists(output_filepath):
        file_mtime = int(os.path.getmtime(output_filepath))
        print(f"既存ファイル '{file_name_only}' のタイムスタンプ: {datetime.datetime.fromtimestamp(file_mtime)}")
    else:
        print(f"対象ファイル '{file_name_only}' は存在しません。新規書き込みを行います。")

    try:
        print(f"✅ データベースに接続しました。")

        # 4. データのクエリ
        record = session.get(Wiki, target_name)

        if record:
            file_content = record.text
            db_time = record.time

            # 5. タイムスタンプの比較
            if db_time and db_time > file_mtime:
                
                print(f"⏳ DBのtime: {datetime.datetime.fromtimestamp(db_time)}")
                print("➡️ **データベースのデータが新しい**ため、ファイルを上書きします。")

                # 6. ファイルへの書き出し（上書き）
                with open(output_filepath, 'w', encoding='utf-8') as f:
                    f.write(file_content)
                
                print(f"✨ ファイル **{output_filepath}** を正常に更新しました。")
                
            else:
                db_time_display = datetime.datetime.fromtimestamp(db_time) if db_time else "N/A"
                print(f"⏳ DBのtime: {db_time_display}")
                print("⏸️ **ファイルのタイムスタンプ以降に更新がない**ため、書き出しをスキップします。")

        else:
            print(f"⚠️ nameが '{target_name}' のレコードは見つかりませんでした。")

    except Exception as error:
        print(f"❌ エラーが発生しました: {error}")
    finally:
        # 7. セッションのクローズ
        session.close()
        print("🔗 データベース接続を閉じました。")

if __name__ == "__main__":

    # 保存先フォルダの確認と作成
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"保存先フォルダ '{OUTPUT_DIR}' を作成しました。")


    # 呼び出し例 1
    FIRST_TARGET = '__MapPage00002754'
    fetch_and_save_if_newer_orm(FIRST_TARGET)
    
    print("\n" + "="*40 + "\n")

    # 呼び出し例 2
    SECOND_TARGET = 'AnotherWikiPage'
    fetch_and_save_if_newer_orm(SECOND_TARGET)
   