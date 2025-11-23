# test_selenium.py

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time   # ← 追加

options = Options()
# options.add_argument("--headless")  # 画面なしで実行

driver = webdriver.Chrome(options=options)
driver.get("https://www.google.com")

time.sleep(15)  # ← 3秒待つ

print(driver.title)  # → "Google" と表示されればOK

driver.quit()

