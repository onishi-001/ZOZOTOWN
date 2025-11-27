# zozo_auto_upload.py

import pandas as pd
import os
import datetime as dt
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

import platform

# ==============================
#  設定
# ==============================
# EXCEL_PATH = "\\160.251.168.21\Zozotown_List\List\ZOZOアップロード指示.xlsx"
EXCEL_PATH = ""
TARGET_SHEET = "作業カレンダー"
TEXT_DIR   = ""
# UPLOAD_DIR = "\\160.251.168.21\Zozotown_List\List_Data\"
UPLOAD_DIR = "\List_Data"

BASIC_USER = "zozotown-60"
BASIC_PASS = "z02o-tenant-0ff1ce"
FORM_USER  = "yasuda.k"
FORM_PASS  = "2aGSOpDiX111111112"
ZOZO_URL   = f"https://{BASIC_USER}:{BASIC_PASS}@to.zozo.jp/"



# ==============================
# Main（拡張OK）
# ==============================
def zozotown_upload_file():
    """
    zozotown_upload_file()
    ├─ is_wsl()                  起動環境により設定値変更
    ├─ read_excel()              Excel読み込み
    ├─ find_upload_file()        テキスト→CSV変換
    ├─ selenium_upload()         アップロード（失敗なら例外）
    ├─ update_excel_result()     Excel更新
    └─ finish()                  後処理
    """
    global EXCEL_PATH, TEXT_DIR   # ← 重要！

    if is_wsl():  # テスト環境（WSL）
        EXCEL_PATH = "/mnt/z/List/ZOZOアップロード指示.xlsx"
        TEXT_DIR     = "/mnt/zozo/text/"
    else:         # 本番環境（VPS Ubuntu）
        EXCEL_PATH   = "/srv/shared_zozo/List/ZOZOアップロード指示.xlsx"
        TEXT_DIR     = "/home/oni190501/data/text/"



    df = read_excel()
    
    upload_file, df = find_upload_file(df)
    if upload_file is None:
        print("🔸 アップロード対象ファイルがありません")
        return
    
    success = selenium_upload(upload_file)
    update_excel_result(df, success)

    print("☑ 全処理終了")


# ==============================
# ① Excel読み込み
# ==============================
def read_excel():
    """
    Excel ファイルを読み込み、指定シートの DataFrame を返す。
    シートが存在しない場合は警告を出して None を返す。
    """

    try:
        # Excel のシート名一覧を取得
        all_sheets = pd.ExcelFile(EXCEL_PATH).sheet_names
        if TARGET_SHEET not in all_sheets:
            print(f"⚠ シート '{TARGET_SHEET}' が Excel ファイルに存在しません")
            return None

        # 指定シートを読み込み　　A4からデータが保存されている
        df = pd.read_excel(EXCEL_PATH, sheet_name=TARGET_SHEET, header=3)
        return df

    except FileNotFoundError:
        print(f"❌ Excel ファイル '{EXCEL_PATH}' が存在しません")
        return None

    except Exception as e:
        print(f"❌ Excel 読み込み時にエラー: {e}")
        return None


# ==============================
# ② テキスト → CSV / 対象ファイル抽出
# ==============================
def find_upload_file(df):
    now = dt.datetime.now()
    upload_file = None

    for index, row in df.iterrows():
        try:
            if pd.to_datetime(row["日時"]) <= now:
                txt_name = f"{row['テキストファイル名']}.txt"
                txt_path = os.path.join(TEXT_DIR, txt_name)

                if os.path.exists(txt_path):
                    # 行数カウント
                    with open(txt_path, "r") as f:
                        line_count = sum(1 for _ in f)
                    df.at[index, "データ行数表示"] = line_count

                    # CSV作成
                    csv_path = os.path.join(UPLOAD_DIR, f"{row['テキストファイル名']}.csv")
                    with open(txt_path, "r") as f_in, open(csv_path, "w") as f_out:
                        for line in f_in:
                            f_out.write(line)

                    df.at[index, "処理結果"] = "アップロード対象"
                    upload_file = csv_path
                else:
                    df.at[index, "処理結果"] = "ファイル無し"

        except Exception as e:
            df.at[index, "処理結果"] = f"エラー: {e}"

    df.to_excel(EXCEL_PATH, index=False)
    return upload_file, df


# ==============================
# ③ Seleniumアップロード
# ==============================
def selenium_upload(upload_file):
    if not upload_file:
        print("アップロード対象なし → 処理終了")
        return False

    print("Selenium開始:", upload_file)

    options = Options()
    options.add_argument("--lang=ja-JP")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)

    try:
        driver.get(ZOZO_URL)
        wait = WebDriverWait(driver, 15)

        user_input = wait.until(EC.presence_of_element_located((By.ID, "UserID")))
        user_input.send_keys(FORM_USER)
        password_input = driver.find_element(By.NAME, "Password")
        password_input.send_keys(FORM_PASS + Keys.ENTER)
        time.sleep(10)

        driver.get("https://to.zozo.jp/to/Advertisement.asp?c=RegistGoodsAd")
        file_input = wait.until(EC.presence_of_element_located((By.NAME, "upfile")))
        file_input.send_keys(upload_file)

        time.sleep(3)
        print("アップロード完了:", upload_file)
        return True

    except Exception as e:
        print("Selenium エラー:", e)
        return False

    finally:
        driver.quit()


# ==============================
# ④ Excelへ結果反映
# ==============================
def update_excel_result(df, success):
    for index, row in df.iterrows():
        if row["処理結果"] == "アップロード対象":
            df.at[index, "処理結果"] = "処理済み" if success else "アップ失敗"

    df.to_excel(EXCEL_PATH, index=False)
    print("Excel更新完了")

# ==============================
#   グローバル設定 (環境別)
# ==============================
def is_wsl():
    # unameのreleaseに "microsoft" が含まれていれば WSL
    return 'microsoft' in platform.uname().release.lower()



if __name__ == "__main__":
    zozotown_upload_file()

