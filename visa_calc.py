import time
import gspread

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from oauth2client.service_account import ServiceAccountCredentials

# 크롬 드라이버 자동 업데이트
from webdriver_manager.chrome import ChromeDriverManager

# 크롬 옵션
chrome_options = Options()
chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# 비자 환율 페이지 열기
driver.implicitly_wait(1)
driver.get('https://www.visakorea.com/support/consumer/travel-support/exchange-rate-calculator.html')

shadow_host = driver.find_element(By.CSS_SELECTOR, '#skipTo > div:nth-child(1) > div > div:nth-child(1) > div.exchangerateCalculator.section > dm-calculator')
shadow_root = shadow_host.shadow_root

# 계산기에 1 입력
input = shadow_root.find_element(By.CSS_SELECTOR, '#input_amount_paid')
input.click()
input.send_keys('1')

# 거래할 통화
currency_to = shadow_root.find_element(By.CSS_SELECTOR, '#autosuggestinput_to')
currency_to.click()
list = shadow_root.find_elements(By.CSS_SELECTOR, '#listbox-to > li')
list[84].click()
#currency_to.send_keys(Keys.ENTER)

startRow = 2
def get_list():
    currency_from = shadow_root.find_element(By.CSS_SELECTOR, '#autosuggestinput_from')
    currency_from.click()
    list = shadow_root.find_elements(By.CSS_SELECTOR, '#listbox-from > li')
    return list

def check_currency():
    list = get_list()
    for i in range(len(list)): 
        list = get_list() # ???????????????
        i += 1
        #print('i : ', i)
        #print('리스트 : ', len(list))
        if(i == len(list)):
            return
        currency_from = shadow_root.find_element(By.CSS_SELECTOR, '#autosuggestinput_from')
        currency_from.click()
        list[i].click()
        convert()


def convert():
    # 변환 버튼 클릭
    btnCalc = shadow_root.find_element(By.CSS_SELECTOR, 'form > div > div > div:nth-child(3) > div > div:nth-child(1) > button')
    btnCalc.click()
    insertSheet()


# 구글 시트 연동
def insertSheet():
    result = shadow_root.find_element(By.CSS_SELECTOR, 'form > div > div.vs-output > div > h2 > strong')
    split_result = result.text.split(maxsplit=1)
    send_result = split_result[0] # 2.33...

    currency = shadow_root.find_element(By.CSS_SELECTOR, 'form > div > div.vs-output > div > h2')
    split_currency = currency.text.split(maxsplit=2)
    send_currency = split_currency[1]

    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
    client = gspread.authorize(creds)
    doc = client.open_by_url('https://docs.google.com/spreadsheets/d/1tt3aCgkLYgsQ04G3ewMi1zKCArFILUUJu0svQOtqm9w/edit#gid=0')
    sheet = doc.worksheet('비자')
    global startRow
    print(send_currency)
    sheet.update_cell(startRow, 2, send_currency)
    print(send_result)
    sheet.update_cell(startRow, 3, send_result)
    print("실행 열 : ", startRow)
    startRow += 1
    time.sleep(1.5)

print("통화 체크")
check_currency()
print("입력 종료")
driver.close()