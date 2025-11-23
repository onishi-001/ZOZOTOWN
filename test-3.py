# zozo_auto_login.py

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# -------------------------
# 設定
# -------------------------
BASIC_USER = "zozotown-60"
BASIC_PASS = "z02o-tenant-0ff1ce"

FORM_USER = "yasuda.k"
FORM_PASS = "2aGSOpDiX111111112"

# Basic認証をURLに埋め込む
ZOZO_URL = f"https://{BASIC_USER}:{BASIC_PASS}@to.zozo.jp/"

# -------------------------
# Chrome / Selenium設定
# -------------------------
options = Options()
options.add_argument("--lang=ja-JP")          # 日本語表示
# options.add_argument("--headless")            # 画面なし
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--font-family=Arial,Meiryo,MS Gothic")  # 日本語フォント指定


driver = webdriver.Chrome(options=options)

try:
    # ページにアクセス（第一認証自動突破）
    driver.get(ZOZO_URL)
    
    # 第二認証フォームが現れるまで待機
    wait = WebDriverWait(driver, 15)
    user_input = wait.until(
        EC.presence_of_element_located((By.ID, "UserID"))
    )
    
    # 第二認証フォーム入力
    user_input.send_keys(FORM_USER)
    password_input = driver.find_element(By.NAME, "Password")
    password_input.send_keys(FORM_PASS + Keys.ENTER)

    # ログイン後ページが開くまで少し待つ
    time.sleep(10)
    
    print("ログイン完了ページタイトル:", driver.title)

    # --- 広告登録ページへ移動 ---
    driver.get("https://to.zozo.jp/to/Advertisement.asp?c=RegistGoodsAd")

    wait = WebDriverWait(driver, 15)

    # input[type='file'] を待機して取得
    file_input = wait.until(
        EC.presence_of_element_located((By.NAME, "upfile"))
    )

    time.sleep(10)

    # アップロードするファイルの絶対パス
    FILE_PATH = "/home/oni190501/zozo_env/upload/ad_sample.csv"  # 例

    # ファイルを選択（クリックは不要）

    # file_input.send_keys(FILE_PATH)

    print("ファイル選択 OK！")



    # -------------------------
    # ログアウト

    # ドロップダウンメニューを開く（必要なら）
    menu = driver.find_element(By.CLASS_NAME, "header-general-menu")
    driver.execute_script("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", menu)

    # JSでクリック
    logout_link = driver.find_element(By.XPATH, "//a[contains(@href,'Default.asp?c=Logout')]")
    driver.execute_script("arguments[0].click();", logout_link)
    time.sleep(10)

    print("ログアウト完了（JSクリック）")

finally:
    driver.quit()

