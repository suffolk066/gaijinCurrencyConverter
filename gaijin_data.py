import time
import random
import gspread

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from oauth2client.service_account import ServiceAccountCredentials

# 크롬 드라이버 자동 업데이트
from webdriver_manager.chrome import ChromeDriverManager

# 크롬 옵션
chrome_options = Options()
user_data = r"C:\Users\zofld\AppData\Local\Google\Chrome\User Data"
chrome_options.add_argument(f"user-data-dir={user_data}")
#chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# 상점 열기
driver.implicitly_wait(3) # 3초 동안 기다림
# driver.maximize_window() # 화면 최대화함
driver.get('https://store.gaijin.net/user.php?skin_lang=en')

# 로그인 버튼 클릭
try:
    if driver.find_element(By.CSS_SELECTOR, "#__container > div.row.js-account-list-container > div:nth-child(1) > ul > li > a").is_displayed():
        print("로그인 버튼 존재")
        btn_login = driver.find_element(By.CSS_SELECTOR, "#__container > div.row.js-account-list-container > div:nth-child(1) > ul > li > a")
        btn_login.click()
        time.sleep(random.uniform(1, 2))
except:
    print("로그인 버튼 없음")

# 가이진 코인 충전 페이지 이동
print("충전 페이지 이동")
btn_balance = driver.find_element(By.CSS_SELECTOR, "#gaijin-menu > nav > ul > li:nth-child(4) > a")
btn_balance.click()
time.sleep(random.uniform(1, 2))

# 1코인 선택
print("1코인 선택")
btn_coin = driver.find_element(By.CSS_SELECTOR, "#bodyRoot > div.content > section > div.balance > div.balance__wrapper > div.balance__refill.balance-refill > ul > li:nth-child(1) > div > div.balance-refill__fix-sum-link.js-fix-refill.js-popup__trigger")
btn_coin.click()
time.sleep(random.uniform(1, 2))

# 결제창 클릭
print("결제창 클릭")
btn_confirm = driver.find_element(By.CSS_SELECTOR, "#g2s_cc_usd")
btn_confirm.click()
driver.implicitly_wait(4)


# 결제창 프레임 전환
print("결제창 대기")
driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="pay-frame"]'))
driver.implicitly_wait(3)

# 통화 선택
select = Select(driver.find_element(By.XPATH, '//*[@id="currency_lbl"]/div/select'))


# 구글 시트 연동
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
client = gspread.authorize(creds)
doc = client.open_by_url('https://docs.google.com/spreadsheets/d/1tt3aCgkLYgsQ04G3ewMi1zKCArFILUUJu0svQOtqm9w/edit#gid=0')
sheet = doc.worksheet('구글')
# sheet.resize(1, 3) # 시트 1행 3열로 만들기




# 통화 데이터 가져오기
start_row = 3
for item in select.options:
    get_inner_text = item.get_attribute('innerText') # n번째 통화 얻어오기
    select.select_by_value(get_inner_text) # n번째 통화 선택
    time.sleep(2) # 통화 선택하고 2초동안 대기 
    # ㄴ 구글 시트 쓰기 요청 분당 60개만 가능하기 때문에 딜레이 길게함

    amount = driver.find_element(By.CSS_SELECTOR, "#orderTotalAmount") # 통화 가격 선택
    currency_amount = ""
    if (amount.text.count(' ') == 1):
        split_amount = amount.text.split(maxsplit=1)
        currency_amount = split_amount[1]
    else:
        split_amount = amount.text.split(maxsplit=2)
        currency_amount = split_amount[1] + split_amount[2]

    currency_text = item.get_attribute('value') # n번째 통화 영문 3글자 따오기 ex) USD, KRW, JPY...

    print(currency_text + ' ' + currency_amount) # 콘솔에 현재 통화 데이터 1줄씩 출력
    # sheet.insert_row([start_row, currency_text, currency_amount], start_row) # 구글 시트에 1행부터 데이터 삽입
    sheet.update_cell(start_row, 3, currency_amount)
    start_row += 1

print("출력 종료")

# 종료
driver.close()