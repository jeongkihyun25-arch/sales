import os, time, shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')
current_dir = os.path.abspath(os.getcwd())
temp_dir = os.path.join(current_dir, 'temp_downloads')
os.makedirs(temp_dir, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_experimental_option("prefs", {"download.default_directory": temp_dir})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': temp_dir})

try:
    print("🚀 WOS 세일즈 실적 수집 (원본 유지 방식)...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)
    
    s_date = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(25) 
    driver.find_element(By.CLASS_NAME, 'excel_btn').click()
    time.sleep(30)

    files = [f for f in os.listdir(temp_dir) if not f.endswith('.crdownload')]
    if files:
        source_path = os.path.join(temp_dir, files[0])
        target_path = os.path.join(current_dir, "current_sales.xlsx")

        print("📊 WOS 다운로드 원본을 진짜 엑셀로만 변환합니다...")
        try:
            new_df = pd.read_excel(source_path, engine='xlrd')
        except:
            with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
                new_df = pd.read_html(f, flavor='html5lib')[0]

        # 💡 [핵심] 원본 모양 그대로 저장 (합성 절대 안 함)
        new_df.to_excel(target_path, index=False)
        print("🎉 current_sales.xlsx 원본 변환 완료 및 저장 성공!")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    raise e
finally:
    driver.quit()
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
