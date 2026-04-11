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

# 1. 시크릿 정보 (GitHub Secrets)
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 및 폴더 설정
current_dir = os.path.abspath(os.getcwd())
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--window-size=1920,1080')

options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': temp_download_dir})

try:
    # 4. 로그인
    print("🚀 WOS 접속 및 로그인 시작...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    wait = WebDriverWait(driver, 20)
    user_id_el = wait.until(EC.presence_of_element_located((By.ID, 'userID')))
    user_id_el.send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    # 5. 매출현황 이동
    print("📂 매출현황 페이지 이동 중...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 6. 날짜 조회 (2026-01-01 고정)
    print("🔍 데이터 조회 버튼 클릭 (2026-01-01 기준)...")
    s_date_input = wait.until(EC.presence_of_element_located((By.ID, 'sDate')))
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(25)

    # 7. 엑셀 다운로드
    print("📥 엑셀 다운로드 버튼 클릭...")
    excel_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'excel_btn')))
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 파일 다운로드 대기
    print("⏳ 파일 생성 대기 중...")
    latest_file = None
    for i in range(60):
        files = [f for f in os.listdir(temp_download_dir) if not f.endswith('.crdownload') and not f.startswith('.')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            break
        time.sleep(2)

    # 9. 파일 변환 및 저장 (한글 인코딩 해결 버전)
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"📊 한글 인코딩(CP949) 적용 변환 시작...")
        try:
            # 💡 WOS 특유의 한글 인코딩 HTML을 안전하게 읽어옵니다.
            with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
                df_list = pd.read_html(f, flavor='html5lib')
            
            if df_list:
                df_list[0].to_excel(target_path, index=False)
                print(f"🎉 성공! current_sales.xlsx 갱신 완료.")
        except Exception as e:
            print(f"⚠️ 1차 변환 실패, 일반 시도: {e}")
            df = pd.read_excel(source_path)
            df.to_excel(target_path, index=False)
    else:
        raise Exception("❌ 파일을 받지 못했습니다.")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    exit(1)
finally:
    driver.quit()
    if os.path.exists(temp_download_dir):
        shutil.rmtree(temp_download_dir)
