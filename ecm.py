# ecm.py
import os
import time
import contextlib
from typing import Optional, List

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

LOGIN_PAGE_URL = "http://210.96.71.85/auth/login/loginView.do"

# =========================
# Chrome 드라이버 생성
# =========================
def _build_chrome(
    headless: bool = True,
    images_enabled: bool = False,
    page_load_strategy: str = "eager",
    page_load_timeout: int = 15,
) -> webdriver.Chrome:
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-notifications")
    if not images_enabled:
        options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.set_capability("pageLoadStrategy", page_load_strategy)

    service = Service(log_path=os.devnull)
    driver = webdriver.Chrome(options=options, service=service)
    driver.set_page_load_timeout(page_load_timeout)
    return driver


# =========================
# (선택) SESSION_KEY만 얻기
# =========================
def get_session_key(
    user_id: str,
    password: str,
    login_url: str = LOGIN_PAGE_URL,
    wait_timeout: float = 6.0,
) -> Optional[str]:
    driver = _build_chrome(headless=True, images_enabled=False)
    try:
        driver.get(login_url)
        wait = WebDriverWait(driver, timeout=wait_timeout, poll_frequency=0.2)
        user_el = wait.until(EC.presence_of_element_located((By.NAME, "user_id")))
        pwd_el  = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        user_el.clear(); user_el.send_keys(user_id)
        pwd_el.clear();  pwd_el.send_keys(password)

        btn = driver.find_element(By.CLASS_NAME, "btn-login")
        driver.execute_script("arguments[0].click();", btn)

        def has_session_cookie(drv):
            ck = drv.get_cookie("SESSION_KEY")
            return ck["value"] if ck and ck.get("value") else False

        skey = WebDriverWait(driver, timeout=5, poll_frequency=0.2).until(has_session_cookie)
        return str(skey) if skey else None
    finally:
        with contextlib.suppress(Exception):
            driver.quit()


# =========================
# 로그인 완료 드라이버 반환
# =========================
def login_and_return_driver(
    user_id: str,
    password: str,
    login_url: str = LOGIN_PAGE_URL,
    headless: bool = True,
) -> webdriver.Chrome:
    driver = _build_chrome(headless=headless, images_enabled=False)
    driver.get(login_url)

    wait = WebDriverWait(driver, timeout=10, poll_frequency=0.2)
    user_el = wait.until(EC.presence_of_element_located((By.NAME, "user_id")))
    pwd_el  = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    user_el.clear(); user_el.send_keys(user_id)
    pwd_el.clear();  pwd_el.send_keys(password)

    btn = driver.find_element(By.CLASS_NAME, "btn-login")
    driver.execute_script("arguments[0].click();", btn)
    time.sleep(0.8)  # 페이지 안정화
    return driver


# -------------------------
# 트리 탐색 유틸 (jsTree 전용)
# -------------------------
class _TreeSelectors:
    # 스크린샷/DOM 기준 고정
    CONTAINER = ["#edm-folder"]
    ITEM_NODE = ["li.jstree-node"]
    TOGGLE    = ["i.jstree-ocl"]
    LABEL     = ["a.jstree-anchor"]


def _detect_selector(driver: webdriver.Chrome, candidates: List[str]) -> str:
    for css in candidates:
        if driver.find_elements(By.CSS_SELECTOR, css):
            return css
    return candidates[0]


def _maybe_switch_to_tree_iframe(driver: webdriver.Chrome) -> None:
    """
    #edm-folder가 보이는 프레임으로 자동 전환.
    """
    # 기본 문서에 있으면 OK
    if driver.find_elements(By.CSS_SELECTOR, "#edm-folder"):
        return

    # iframe 순회
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for f in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(f)
            if driver.find_elements(By.CSS_SELECTOR, "#edm-folder"):
                return
        except Exception:
            continue

    # 못 찾으면 기본 문서로 복귀
    driver.switch_to.default_content()


def _click_menu_if_needed(driver: webdriver.Chrome) -> None:
    """
    로그인 직후 트리가 안 뜨면 '문서관리' 메뉴 클릭 시도.
    """
    if driver.find_elements(By.CSS_SELECTOR, "#edm-folder"):
        return

    candidates = [
        (By.LINK_TEXT, "문서관리"),
        (By.PARTIAL_LINK_TEXT, "문서관리"),
        (By.XPATH, "//a[contains(., '문서관리') or contains(@href,'edm')]"),
        (By.XPATH, "//button[contains(., '문서관리')]"),
    ]
    for how, sel in candidates:
        try:
            el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((how, sel)))
            driver.execute_script("arguments[0].click();", el)
            time.sleep(0.5)
            if driver.find_elements(By.CSS_SELECTOR, "#edm-folder"):
                return
        except Exception:
            pass


def _is_open(node) -> bool:
    cls = (node.get_attribute("class") or "").lower()
    return "jstree-open" in cls


def _is_closed(node) -> bool:
    cls = (node.get_attribute("class") or "").lower()
    return "jstree-closed" in cls


def _expand_node(driver: webdriver.Chrome, node) -> None:
    """
    닫힌 노드면 토글(i.jstree-ocl) 클릭하여 펼침.
    """
    if not _is_closed(node):
        return
    try:
        toggle = node.find_element(By.CSS_SELECTOR, "i.jstree-ocl")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", toggle)
        try:
            toggle.click()
        except Exception:
            driver.execute_script("arguments[0].click();", toggle)

        # 열린 상태 대기
        for _ in range(20):
            if _is_open(node):
                break
            time.sleep(0.05)
    except Exception:
        pass


def _label_text(node) -> str:
    try:
        a = node.find_element(By.CSS_SELECTOR, "a.jstree-anchor")
        return (a.text or "").strip()
    except Exception:
        return (node.text or "").strip()


def _build_full_path(node) -> str:
    """
    li 기준으로 조상 li의 라벨까지 모아 경로 생성.
    """
    labels = []
    cur = node
    for _ in range(80):
        if cur.tag_name.lower() == "li":
            t = _label_text(cur)
            if t:
                labels.append(t)
        try:
            parent = cur.find_element(By.XPATH, "./..")
        except Exception:
            break
        if parent is None or parent == cur:
            break
        cur = parent
        # li가 아니면 한 단계 더 상위로
        try:
            cur.find_element(By.XPATH, "self::li")
        except Exception:
            try:
                cur = cur.find_element(By.XPATH, "./..")
            except Exception:
                break
    labels.reverse()
    return "/".join([x for x in labels if x])


def _find_root_node(driver: webdriver.Chrome, root_label: str):
    """
    앵커 텍스트가 root_label인 li 노드를 반환.
    """
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.jstree-anchor")
    for a in anchors:
        if (a.text or "").strip() == root_label:
            try:
                return a.find_element(By.XPATH, "./ancestor::li[1]")
            except Exception:
                return None
    return None


def _relativize(full_path: str, root_label: str, include_root: bool = True) -> Optional[str]:
    parts = [p for p in full_path.split("/") if p]
    if root_label not in parts:
        return None
    i = parts.index(root_label)
    rel = parts[i:] if include_root else parts[i + 1 :]
    # 루트 자신만은 제외
    return "/".join(rel) if "/".join(rel) != root_label else None


# =========================
# 핵심: 특정 루트 하위만 수집
# =========================
def collect_paths_under_root(
    user_id: str,
    password: str,
    root_label: str = "영남소프트웨어시험센터",
    login_url: str = LOGIN_PAGE_URL,
    headless: bool = True,
    include_root_in_path: bool = True,
) -> List[str]:
    """
    로그인 후, 좌측 트리에서 `root_label` 노드를 찾아 전개하고
    그 하위 모든 폴더의 '전체 경로'(기본: 루트 포함)를 리스트로 반환.
    """
    driver = login_and_return_driver(user_id, password, login_url=login_url, headless=headless)
    try:
        # 트리 화면 보장
        _click_menu_if_needed(driver)
        _maybe_switch_to_tree_iframe(driver)

        # 컨테이너 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#edm-folder")))

        # 루트 노드 찾기
        root_node = _find_root_node(driver, root_label)
        if root_node is None:
            # 디버깅 도움: 상위 앵커 텍스트 샘플
            samples = [a.text.strip() for a in driver.find_elements(By.CSS_SELECTOR, "a.jstree-anchor")[:30]]
            raise RuntimeError(f"루트 노드 '{root_label}'를 찾지 못했습니다. 앵커 샘플: {samples}")

        # BFS로 서브트리 모두 펼치기
        queue = [root_node]
        seen = set()
        while queue:
            node = queue.pop(0)
            node_key = node.get_attribute("id") or node.get_attribute("oid") or str(node)
            if node_key in seen:
                continue
            seen.add(node_key)

            _expand_node(driver, node)

            try:
                children = node.find_elements(By.XPATH, "./ul/li")
                queue.extend(children)
            except Exception:
                pass

        # 경로 수집 (루트 자신 제외)
        all_nodes = [root_node] + root_node.find_elements(By.XPATH, ".//li")
        result: List[str] = []
        for n in all_nodes:
            p = _build_full_path(n)
            if not p or p == root_label:
                continue
            rel = _relativize(p, root_label, include_root=include_root_in_path)
            if rel:
                result.append(rel)

        return sorted(set(result))

    finally:
        with contextlib.suppress(Exception):
            driver.quit()


# =========================
# 수동 실행 테스트
# =========================
if __name__ == "__main__":
    uid = os.getenv("ECM_USER") or "jeonje"
    pwd = os.getenv("ECM_PASS") or "qwasqwas12"

    paths = collect_paths_under_root(
        uid, pwd,
        root_label="영남소프트웨어시험센터",
        headless=False,                 # 디버깅 시 False
        include_root_in_path=True       # "영남소프트웨어시험센터/...." 형태
    )
    print(f"[총 {len(paths)}개]")
    for p in paths[:80]:
        print(p)
