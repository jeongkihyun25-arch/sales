import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

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
    print("🚀 WOS 접속 및 로그인...")
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
            print(f"✅ 다운로드 완료: {latest_file}")
            break
        time.sleep(2)

    # 6. 💡 [핵심] LT 때 썼던 웹 형식 변환 비법 적용
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"📊 데이터 변환 공정 시작...")
        
        # [비법 1] 웹 형식(HTML)으로 먼저 시도 (대부분의 WOS 파일)
        try:
            print("...1단계 시도: 웹 형식(HTML/CP949) 읽기")
            with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
                df_list = pd.read_html(f, flavor='html5lib')
            if df_list:
                df_list[0].to_excel(target_path, index=False)
                print(f"🎉 성공! 웹 형식을 진짜 엑셀로 변환했습니다.")
            else:
                raise Exception("표를 찾을 수 없음")
        
        # [비법 2] 웹 형식이 아닐 경우, 진짜 옛날 엑셀(.xls)로 시도
        except Exception as e:
            print(f"⚠️ 1단계 실패, 2단계 시도: 진짜 옛날 엑셀(Binary XLS) 읽기")
            try:
                # 여기서 xlrd 엔진을 사용하여 읽습니다.
                df = pd.read_excel(source_path, engine='xlrd')
                df.to_excel(target_path, index=False)
                print(f"✅ 성공! 옛날 엑셀 형식을 최신 엑셀로 변환했습니다.")
            except Exception as e2:
                print(f"❌ 모든 시도 실패. 원본 파일 자체에 문제가 있습니다: {e2}")
                # 최후의 수단: 그냥 복사라도 합니다 (이때 엑셀 에러가 날 수 있음)
                shutil.copy2(source_path, target_path)
    else:
        raise Exception("❌ 파일을 다운로드하지 못했습니다.")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    exit(1)
finally:
    driver.quit()
    if os.path.exists(temp_download_dir):
        shutil.rmtree(temp_download_dir)
