import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 1. 시크릿 정보
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 설정
current_dir = os.path.abspath(os.getcwd())
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 옵션
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': temp_download_dir})

try:
    # 4. 로그인 및 페이지 이동
    print("🚀 WOS 접속...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    print("📂 매출현황 이동 및 조회...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 2026-01-01부터 조회
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(25) 

    print("📥 엑셀 다운로드 시도...")
    driver.find_element(By.CLASS_NAME, 'excel_btn').click()
    
    # 5. 파일 대기
    latest_file = None
    for i in range(50):
        files = [f for f in os.listdir(temp_download_dir) if not f.endswith('.crdownload')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            print(f"✅ 다운로드 완료 확인: {latest_file}")
            break
        time.sleep(2)

    # 6. 💡 [LT 대시보드 비법] 웹 형식(HTML) 강제 변환 파트
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"📊 [핵심] 웹 코드를 진짜 엑셀로 세탁 중...")
        try:
            # WOS 파일은 사실 '가짜 엑셀'입니다. 텍스트로 읽어서 HTML 테이블을 뽑아내야 합니다.
            # 인코딩은 한글 깨짐 방지를 위해 cp949를 사용합니다.
            with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
                html_text = f.read()
                # pandas의 read_html을 사용하여 웹 표 정보를 긁어옵니다.
                df_list = pd.read_html(html_text, flavor='html5lib')
            
            if df_list:
                df = df_list[0]
                # 여기서 진짜 엑셀(.xlsx) 파일로 저장합니다.
                df.to_excel(target_path, index=False)
                print(f"🎉 성공! 웹 형식을 완벽하게 엑셀로 변환했습니다.")
            else:
                raise Exception("데이터 테이블을 찾을 수 없습니다.")
                
        except Exception as e:
            print(f"⚠️ 웹 형식 변환 실패, 최후의 수단(강제 복사) 시도: {e}")
            shutil.copy2(source_path, target_path)
    else:
        raise Exception("❌ 결국 파일을 받지 못했습니다.")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    exit(1)
finally:
    driver.quit()
    if os.path.exists(temp_download_dir):
        shutil.rmtree(temp_download_dir)
