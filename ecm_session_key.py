import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

LOGIN_PAGE_URL = "http://210.96.71.85/auth/login/loginView.do"
LOGIN_DATA = {
    "user_id": "jeonje",
    "password": "qwasqwas12",
}

# ---- Chrome 옵션 (속도/로그 최적화) ----
options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")                      # Windows에서도 무해, 컨텍스트 단순화
options.add_argument("--disable-dev-shm-usage")           # 공유 메모리 이슈 완화
options.add_argument("--disable-extensions")
options.add_argument("--disable-background-networking")
options.add_argument("--disable-default-apps")
options.add_argument("--disable-notifications")
options.add_argument("--blink-settings=imagesEnabled=false")  # 이미지 로딩 차단
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 콘솔 로그 억제
# 페이지 로드 전략: DOMContentLoaded까지만 대기
options.set_capability("pageLoadStrategy", "eager")

# Chromedriver 자체 로그 버리기(속도엔 큰 영향 없지만 콘솔 깔끔)
service = Service(log_path=os.devnull)

driver = webdriver.Chrome(options=options, service=service)
driver.set_page_load_timeout(15)  # 무한 대기 방지

try:
    driver.get(LOGIN_PAGE_URL)

    # 대기 설정: 짧고 자주 체크
    wait = WebDriverWait(driver, timeout=6, poll_frequency=0.2)

    # 입력 필드 로드 후 즉시 입력
    user_el = wait.until(EC.presence_of_element_located((By.NAME, "user_id")))
    pwd_el  = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    user_el.clear(); user_el.send_keys(LOGIN_DATA["user_id"])
    pwd_el.clear();  pwd_el.send_keys(LOGIN_DATA["password"])

    # 버튼 클릭: 렌더 지연 회피용 JS 클릭(기본 클릭보다 빠른 경우 있음)
    btn = driver.find_element(By.CLASS_NAME, "btn-login")
    driver.execute_script("arguments[0].click();", btn)

    # 페이지 전환을 길게 기다리지 않고, 세션 쿠키가 생기는지만 체크
    # (일반적으로 인증 성공 시 쿠키가 즉시 설정됨)
    def has_session_cookie(drv):
        ck = drv.get_cookie("SESSION_KEY")
        return ck["value"] if ck and ck.get("value") else False

    session_key = WebDriverWait(driver, timeout=5, poll_frequency=0.2).until(has_session_cookie)

finally:
    driver.quit()

if session_key:
    print(f"session_key: {session_key}")
else:
    print("SESSION_KEY 없음")
