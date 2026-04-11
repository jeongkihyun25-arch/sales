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
    print("🚀 WOS 세일즈 실적 수집 시작...")
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
        ref_path = os.path.join(current_dir, "historical_data.xlsx")
        target_path = os.path.join(current_dir, "current_sales.xlsx")

        print(f"📊 다운로드 완료. 데이터 합성 시작...")

        # 💡 [핵심 수정] 에러 안 나게 읽는 방법 2중 방어망 구축
        try:
            print("... 1단계: 진짜 엑셀(xlrd) 방식으로 읽기 시도")
            new_df = pd.read_excel(source_path, engine='xlrd')
        except Exception as e1:
            print(f"⚠️ 엑셀 방식 실패, 2단계: 웹(HTML) 방식으로 읽기 시도: {e1}")
            with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
                new_df = pd.read_html(f, flavor='html5lib')[0]

        print("✅ 데이터 읽기 성공! 매핑을 시작합니다.")

        # 기준표(Historical) 읽기
        ref_df = pd.read_excel(ref_path)

        # 실적 합산 및 공백 제거 (3중 매칭)
        new_agg = new_df.groupby(['거래처', 'SIZE', 'PTTN'])['합계수량'].sum().reset_index()
        new_agg.columns = ['거래처명', '사이즈', '패턴명', '실적']
        
        for df, cols in [(ref_df, ['거래처명', '사이즈', '패턴명']), (new_agg, ['거래처명', '사이즈', '패턴명'])]:
            for col in cols:
                df[col] = df[col].astype(str).str.strip()

        # 데이터 끼워넣기
        merged = pd.merge(ref_df, new_agg, on=['거래처명', '사이즈', '패턴명'], how='left')
        ref_df['2026년(당해)'] = merged['실적'].fillna(0)
        
        # 최종 완성본 저장
        ref_df.to_excel(target_path, index=False)
        print("🎉 합성 완료! 완벽한 current_sales.xlsx가 생성되었습니다.")

except Exception as e:
    print(f"❌ 치명적 에러 발생: {e}")
    raise e
finally:
    driver.quit()
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
