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

# 1. 시크릿 정보 (GitHub Secrets에서 가져옴)
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 설정 (깔끔한 관리를 위해 임시 폴더 사용)
current_dir = os.path.abspath(os.getcwd())
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

# 실행 전 기존 임시 폴더 삭제 후 재생성
if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')

# 다운로드 자동 승인 및 경로 지정
options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 💡 중요: 깃허브 가상 서버(Headless)에서 다운로드를 강제 허용하는 명령
driver.execute_cdp_cmd('Page.setDownloadBehavior', {
    'behavior': 'allow',
    'downloadPath': temp_download_dir
})

try:
    # 4. 로그인 과정
    print("🚀 WOS 접속 및 로그인 시작...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    
    wait = WebDriverWait(driver, 20)
    user_id_el = wait.until(EC.presence_of_element_located((By.ID, 'userID')))
    user_id_el.send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    
    # 로그인 처리 대기
    time.sleep(10) 

    # 5. 매출현황 페이지 이동
    print("📂 매출현황 페이지 이동 중...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 6. 날짜 조회 (2026-01-01 고정)
    print("🔍 데이터 조회 버튼 클릭 (2026-01-01 기준)...")
    s_date_input = wait.until(EC.presence_of_element_located((By.ID, 'sDate')))
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    
    search_btn = driver.find_element(By.CLASS_NAME, 'search_btn')
    driver.execute_script("arguments[0].click();", search_btn)
    time.sleep(25) # 데이터 양이 많으므로 충분히 대기

    # 7. 엑셀 다운로드 클릭
    print("📥 엑셀 다운로드 버튼 클릭 시도...")
    excel_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'excel_btn')))
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 파일 생성 확인 (루프 돌며 대기)
    print("⏳ 파일 다운로드 대기 중...")
    latest_file = None
    for i in range(60): # 최대 120초 대기
        files = [f for f in os.listdir(temp_download_dir) 
                 if not f.endswith('.crdownload') and not f.startswith('.')]
        
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            print(f"✅ 다운로드 확인됨: {latest_file}")
            break
        
        if i % 5 == 0:
            print(f"... 기다리는 중 (현재 폴더: {os.listdir(temp_download_dir)})")
        time.sleep(2)

    # 9. 파일 변환 및 저장 (핵심!)
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        # 💡 요청하신 대로 current_sales.xlsx 라는 이름으로 저장합니다.
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"📊 가짜 엑셀(HTML)을 진짜 엑셀(.xlsx)로 변환 중...")
        try:
            # WOS 엑셀은 형식이 HTML이므로 lxml 또는 html5lib로 읽어야 안전합니다.
            df_list = pd.read_html(source_path, flavor='html5lib')
            if df_list:
                df = df_list[0]
                # 변환하여 저장
                df.to_excel(target_path, index=False)
                print(f"🎉 성공! current_sales.xlsx 파일이 갱신되었습니다.")
        except Exception as e:
            print(f"⚠️ 변환 실패, 원본 강제 복사 시도: {e}")
            shutil.copy2(source_path, target_path)
    else:
        raise Exception("❌ 파일을 찾을 수 없습니다. 다운로드에 실패했거나 서버 응답이 너무 느립니다.")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    exit(1)
finally:
    driver.quit()
    # 작업 완료 후 임시 폴더 삭제
    if os.path.exists(temp_download_dir):
        shutil.rmtree(temp_download_dir)
