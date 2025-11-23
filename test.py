# zozo_login_test.py

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

ZOZO_URL = "https://to.zozo.jp/"  # LOGIN画面
USERNAME = "zozotown-60"
PASSWORD = "z02o-tenant-0ff1ce"

options = Options()
# options.add_argument("--headless")  # VPS 移行時にそのまま使える
driver = webdriver.Chrome(options=options)

driver.get(ZOZO_URL)

# --- ログインフォーム操作例 ---
driver.find_element(By.NAME, "login_id").send_keys(USERNAME)
driver.find_element(By.NAME, "password").send_keys(PASSWORD + Keys.ENTER)

time.sleep(3)  # 必要なら調整

print("Page Title:", driver.title)  # ログイン成功確認用
driver.quit()

