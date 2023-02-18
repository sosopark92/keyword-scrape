# pip install selenium
# pip install webdriver_manager


from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

# 엑셀파일 열기
wb = Workbook()
ws = wb.create_sheet('22_과일_쇼핑검색TOP500')
wb.remove_sheet(wb['Sheet'])
ws.append((['월', '순위', '인기검색어']))

chrome_options = webdriver.ChromeOptions()
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
chrome_options.add_argument('user-agent=' + user_agent)
driver = webdriver.Chrome(service=Service(
    ChromeDriverManager().install()), options=chrome_options)

# 네이버 데이터랩 창 열기
driver.get("https://datalab.naver.com/shoppingInsight/sCategory.naver")
time.sleep(5)

# 기기별 모바일 //*[@id="18_device_2"]
# 성별 전체 //*[@id="19_gender_0"]
# 연령 304050 //*[@id="20_age_3"]
# 조회하기 //*[@id="content"]/div[2]/div/div[1]/div/a
# 다음버튼 //*[@id="content"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/a[2]

# driver.find_element(
#     By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[3]/ul/li[4]/a').click()  # 채소
# driver.find_element(
#     By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[3]/ul/li[1]/a').click()  # 쌀
# driver.find_element(
#     By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[3]/ul/li[2]/a').click()  # 잡곡/혼합곡


driver.find_element(By.XPATH, '//*[@id="18_device_2"]').click()
time.sleep(1)  # 기기 모바일
driver.find_element(By.XPATH, '//*[@id="19_gender_0"]').click()  # 성별 전체
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="20_age_3"]').click()  # 연령 30대
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="20_age_4"]').click()  # 연령 40대
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="20_age_5"]').click()  # 연령 50대
time.sleep(1)


# # 분야선택 (식품>농산물>과일)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[1]/span').click()  # 분야
time.sleep(0.5)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[1]/ul/li[7]/a').click()  # 식품
time.sleep(1)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[2]/span').click()  # 2분류
time.sleep(1)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[2]/ul/li[3]/a').click()  # 농산물
time.sleep(0.5)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[3]/span').click()  # 3분류
time.sleep(1)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[1]/div/div[3]/ul/li[3]/a').click()  # 과일
time.sleep(0.5)

# 기간 입력
# 월간 검색 22년 1월
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[1]/div/span').click()
time.sleep(1)
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[1]/div/ul/li[3]/a').click()
time.sleep(1)
# 직접입력
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[1]/span/label[4]').click()
time.sleep(0.5)
# 시작 년도
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[1]/span').click()
time.sleep(0.5)
# 2022년
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[1]/ul/li[6]/a').click()
time.sleep(1)
# 시작 월
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[2]/span').click()
time.sleep(1)
# 1월
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[2]/ul/li[1]/a').click()
time.sleep(2)
# 종료 년도
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[1]/span').click()
time.sleep(1)
# 2022년
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[1]/ul/li[1]/a').click()
time.sleep(2)
# 종료 월
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[2]/span').click()
time.sleep(0.5)
# 1월
driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[2]/ul/li[1]/a').click()
time.sleep(1)


driver.find_element(
    By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/a').click()  # 조회하기

rand_value = randint(5, MAX_SLEEP_TIME)
time.sleep(rand_value)

# 25, 21
for i in range(0, 25):
    for j in range(1, 21):
        path = f'//*[@id="content"]/div[2]/div/div[2]/div[2]/div/div/div[1]/ul/li[{j}]'
        result = "\n1월" + driver.find_element(By.XPATH, path).text
        print(result.split('\n'))
        time.sleep(0.3)
        ws.append(result.split('\n'))

    driver.find_element(
        By.XPATH, '//*[@id="content"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/a[2]').click()
    time.sleep(rand_value)


# 기간 업데이트

# 12, 13
for x in range(0, 12):
    for k in range(2, 13):
        # 종료 월
        driver.find_element(
            By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[2]/span').click()
        time.sleep(0.5)
        # 2월~12월
        driver.find_element(
            By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[3]/div[2]/ul/li[2]/a').click()
        time.sleep(1)
        # 시작 월
        driver.find_element(
            By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[2]/span').click()
        time.sleep(1)
        # 2월~12월
        driver.find_element(
            By.XPATH, f'//*[@id="content"]/div[2]/div/div[1]/div/div/div[2]/div[2]/span[1]/div[2]/ul/li[{k}]/a').click()
        time.sleep(2)

        driver.find_element(
            By.XPATH, '//*[@id="content"]/div[2]/div/div[1]/div/a').click()  # 조회하기
# 25,21
        for i in range(0, 25):
            for j in range(1, 21):
                path = f'//*[@id="content"]/div[2]/div/div[2]/div[2]/div/div/div[1]/ul/li[{j}]'
                result = f"\n{k}월" + driver.find_element(By.XPATH, path).text
                print(result.split('\n'))
                time.sleep(0.3)

            driver.find_element(
                By.XPATH, '//*[@id="content"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/a[2]').click()
            time.sleep(rand_value)
            ws.append(result.split('\n'))

wb.save(r'C:\Users\sosop\OneDrive\바탕 화면\NaverDataLabTop500.xlsx')
wb.close
