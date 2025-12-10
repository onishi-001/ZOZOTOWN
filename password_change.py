# password_change.py

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

from selenium.webdriver.chrome.service import Service   # 自動Download
from selenium.webdriver.chrome.options import Options   # 自動Download
from webdriver_manager.chrome import ChromeDriverManager    # 自動Download

import platform
from openpyxl import load_workbook
from openpyxl.styles import Font

import re

from datetime import datetime

import requests
import json


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
FORM_USER  = "yasuda.k2"
FORM_OLD_PASS  = ""
FORM_NEW_PASS  = ""
FORM_PASS_BASE  = "ADVewV9Bw1"
PASSWORD_FILE = ""

ZOZO_URL   = f"https://{BASIC_USER}:{BASIC_PASS}@to.zozo.jp/"

# 長期トークン LINE
ACCESS_TOKEN_onishi = "5/OgaSQMXxP2DZJg6t7sTFSLlolggNd2zPjsWKd5xjosuYUXuudj7I8KZmZNukWd5jmC5P9+wk6MSojM00MhUrWisjCaufOT0nnf3+K18oixTx7C77I8YydA/0TPCRCx7lDQK9Y48zrpNIoIol+r5wdB04t89/1O/w1cDnyilFU="
ACCESS_TOKEN = "HsMt2pdiuH7B1CpoKcmMoYBLP+Xu4rlURkM+jF5Z8IeFAPQNrKQB+M6KyvYFzWiAhvQWm3NDHi7GinjA3ZrGZFjRsoq7DTPPvPPZFOe+cx8NSG2npKFU/UsKBICIC+JNkY+WXys5w8x25MDMRSvvZgdB04t89/1O/w1cDnyilFU="

# 送信先ユーザーID LINE
USER_ID_onishi = "U615273c685475d75e9d789225d59cb5e"       # onishi
USER_ID = "Ucb0869a3859ac49ac235e0d9efb6bc41"       # yasuda













# ZOZO ログイン URL
ZOZO_LOGIN_URL = "https://to.zozo.jp/to/Default.asp"


def generate_password(length=12):
    chars = string.ascii_letters + string.digits + "!@#$%&*"
    return ''.join(secrets.choice(chars) for _ in range(length))

def change_password():
    """
    change_password()
    ├─ is_wsl()                  起動環境により設定値変更
    ├─ load_password()           パスワードをファイルから取得
    ├─ put_password()            パスワードをファイルに保存する
    ├─ write_log()               ログ出力を行う
    ├─ print_type()              ログを日付付きでprint出力を行う
    └─ line_message()            特定のアカウントにラインメッセージを送信する 400通/月
    """
    
    global PASSWORD_FILE, FORM_OLD_PASS, FORM_NEW_PASS, FORM_PASS_BASE
    global LOG_FILE
    global Error_flag


    # 日付を YYYYMMDD形式で取得
    today_str = datetime.now().strftime("%Y%m%d")

    if is_wsl():  # テスト環境（WSL）
        PASSWORD_FILE = "/mnt/z/Init/Password.txt"
        LOG_FILE = "/mnt/z/Log/" + f"{today_str}.txt"
    else:         # 本番環境（VPS Ubuntu）
        PASSWORD_FILE = "/srv/shared_zozo/Init/Password.txt"
        LOG_FILE = "/srv/shared_zozo/Log/" + f"{today_str}.txt"

    FORM_OLD_PASS = load_password(PASSWORD_FILE)    # 現パスワードを取得する
    FORM_NEW_PASS = FORM_PASS_BASE + "-" + today_str    # 新パスワードを作成する

    print_type("PASSWORD Update開始:")

    options = Options()
    options.add_argument("--lang=ja-JP")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    if not is_wsl():  # 本番環境（VPS Ubuntu）
        options.add_argument("--headless")

    # ChromeDriver を自動で取得
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    wait = WebDriverWait(driver, 10)

    try:
        print_type("STEP 1: ログインページにアクセス")
        driver.get(ZOZO_URL)
        wait = WebDriverWait(driver, 10)


        # --- 1. 最初の画面で「変更」をクリック（ユーザ入力しない） ---
        change_link = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "変更"))
        )
        change_link.click()

        # --- 2. パスワード変更フォームの入力欄の待機 ---
        login_name = wait.until(
            EC.presence_of_element_located((By.NAME, "LoginName"))
        )
        login_name.send_keys(FORM_USER)

        old_pass = driver.find_element(By.NAME, "Password")
        old_pass.send_keys(FORM_OLD_PASS)

        new_pass1 = driver.find_element(By.NAME, "NewPass1")
        new_pass1.send_keys(FORM_NEW_PASS)

        new_pass2 = driver.find_element(By.NAME, "NewPass2")
        new_pass2.send_keys(FORM_NEW_PASS)

        wait = WebDriverWait(driver, 5)

        # --- 3. パスワード変更ボタン押下 ---
        change_btn = driver.find_element(By.NAME, "ChangeBtn")
        change_btn.click()

        put_password(PASSWORD_FILE,FORM_NEW_PASS)    # 現パスワードを取得する

        print_type(f"パスワード変更 成功: {FORM_NEW_PASS}")
        write_log(f"パスワード変更 成功: {FORM_NEW_PASS}")
        line_message(f"パスワード変更 成功: {FORM_NEW_PASS}")

    except Exception as e:
        print_type(f"パスワード変更エラー: {e}")
        write_log(f"パスワード変更エラー: {e}")
        line_message("パスワード変更エラー")
    finally:
        driver.quit()



# ==============================
#   グローバル設定 (環境別)
# ==============================
def is_wsl():
    # unameのreleaseに "microsoft" が含まれていれば WSL
    return 'microsoft' in platform.uname().release.lower()


# ==============================
#   パスワード取得 (環境別)
# ==============================
def load_password(file_path):
    """テキストファイルからパスワードを1行読み込む"""
    with open(file_path, "r", encoding="utf-8") as f:
        password = f.readline().strip()  # 改行を除く
    return password

# ==============================
#   パスワード書込み (環境別)
# ==============================
def put_password(file_path,new_pass):
    """テキストファイルにパスワードを1行書き込む"""
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(new_pass)  # 1行書込む



# ==============================
#   ログ出力 (環境別)
# ==============================

def write_log(message):
    global LOG_FILE
    
    try:
        now = datetime.now().strftime("%H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{now}] -- {message}\n")

    except Exception as e:
        print_type(f"ログ書き込みエラー:{e}")

# ==============================
#   ログ出力 (環境別)
# ==============================

def print_type(message):
    
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] -- {message}\n")

# ==============================
#   LINE メッセージ出力
# ==============================
def line_message(message):
    global ACCESS_TOKEN,USER_ID
    global ACCESS_TOKEN_onishi,USER_ID_onishi

    # メッセージ
    payload = {
        "to": USER_ID,
        "messages": [
            {"type": "text", "text": message}
        ]
    }

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    payload_onishi = {
        "to": USER_ID_onishi,
        "messages": [
            {"type": "text", "text": message}
        ]
    }

    headers_onishi = {
        "Authorization": f"Bearer {ACCESS_TOKEN_onishi}",
        "Content-Type": "application/json"
    }


    url = "https://api.line.me/v2/bot/message/push"


    response = requests.post(url, headers=headers, data=json.dumps(payload))

    print(response.status_code)
    print(response.text)

    # テスト
    response = requests.post(url, headers=headers_onishi, data=json.dumps(payload_onishi))

    print(response.status_code)
    print(response.text)




if __name__ == "__main__":
    change_password()

