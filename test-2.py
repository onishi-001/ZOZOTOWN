from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

ZOZO_URL = "https://to.zozo.jp/"
USERNAME = "zozotown-60"
PASSWORD = "z02o-tenant-0ff1ce"

options = Options()
# options.add_argument("--headless")  # ← コメントアウトするとブラウザ表示される

driver = webdriver.Chrome(options=options)
driver.get(ZOZO_URL)

# JSで生成されるログインフォームを最大15秒待機
login_input = WebDriverWait(driver, 15).until(
    EC.presence_of_element_located((By.NAME, "login_id"))
)
password_input = driver.find_element(By.NAME, "password")

# 第一認証情報を入力
login_input.send_keys(USERNAME)
password_input.send_keys(PASSWORD)
password_input.submit()

# ページ遷移後のタイトル確認
WebDriverWait(driver, 10).until(lambda d: d.title != "")
print("Logged in, page title:", driver.title)

# 確認後に手動で閉じたい場合はコメントアウト
# driver.quit()

