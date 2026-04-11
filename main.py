import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': os.getcwd()})

try:
    print("🚀 WOS 로그인...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10)

    print("📂 매출현황 이동 및 조회...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)
    
    # 1월 1일부터 조회하도록 날짜 설정
    s_date = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(20)

    print("📥 엑셀 다운로드 중...")
    driver.find_element(By.CLASS_NAME, 'excel_btn').click()
    time.sleep(30) # 파일 생성 대기

    # 다운로드된 파일 찾기 및 이름 변경
    files = [f for f in os.listdir('.') if f.endswith('.xls') or f.endswith('.xlsx')]
    if files:
        # 가장 최근 파일을 historical_data.xlsx로 변환
        df_list = pd.read_html(files[0], flavor='html5lib')
        if df_list:
            df_list[0].to_excel('historical_data.xlsx', index=False)
            print("✅ 데이터 업데이트 완료")
finally:
    driver.quit()
