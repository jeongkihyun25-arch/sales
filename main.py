import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 시크릿 정보 가져오기
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
# 다운로드 경로 설정
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
    
    # 2026년 1월 1일부터 오늘까지 조회하도록 날짜 설정
    s_date = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(20)

    print("📥 엑셀 다운로드 중...")
    driver.find_element(By.CLASS_NAME, 'excel_btn').click()
    time.sleep(30) # 파일 생성 및 다운로드 대기

    # 3. 다운로드된 원본 파일 찾기 (WOS는 보통 .xls로 내려줍니다)
    files = [f for f in os.listdir('.') if f.endswith('.xls') or f.endswith('.xlsx')]
    # 기존에 저장된 'current_sales.xlsx'와 'historical_data.xlsx'는 제외하고 찾기
    temp_files = [f for f in files if f not in ['current_sales.xlsx', 'historical_data.xlsx']]

    if temp_files:
        # 가장 최근에 다운로드된 파일 하나 선택
        temp_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        target_file = temp_files[0]
        
        print(f"Parsing: {target_file}")
        # WOS 특유의 HTML형식 엑셀을 읽어오기
        df_list = pd.read_html(target_file, flavor='html5lib')
        
        if df_list:
            # 💡 수정 포인트: 파일명을 current_sales.xlsx로 저장
            df_list[0].to_excel('current_sales.xlsx', index=False)
            print("✅ current_sales.xlsx 업데이트 완료!")
            
            # 작업 끝난 임시 파일 삭제 (다음 실행 때 안 겹치게)
            os.remove(target_file)
            print(f"🗑️ 임시 파일({target_file}) 삭제 완료")
    else:
        print("❌ 다운로드된 새 파일을 찾지 못했습니다.")

finally:
    driver.quit()
