# alzip_zip_capture.py
import os
import sys
import time
import subprocess
from typing import List, Optional, Dict, Tuple
import pyautogui as gui
import pygetwindow as gw

# ====== 설정(기본값) ======
DEFAULT_OPEN_WAIT = 2.2    # ALZip이 열릴 때 대기 (초)
DEFAULT_NAV_WAIT  = 1.0    # Enter 등 이후 안정 대기 (초)
DEFAULT_PRESS_DELAY = 0.08 # 방향키 연속 입력시 각 키 사이 딜레이 (초)
RETRY             = 3

# ---------------- 유틸 ----------------
def _script_parent_dir() -> str:
    p = os.path.abspath(sys.argv[0]); d = os.path.dirname(p)
    return os.path.dirname(d)

def _resolve_out_dir(zip_path: str, out_base_dir: Optional[str]) -> str:
    base = os.path.splitext(os.path.basename(zip_path))[0]
    parent = os.path.abspath(out_base_dir) if out_base_dir else _script_parent_dir()
    out_dir = os.path.join(parent, base)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir

def _sanitize(name: str) -> str:
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, "_")
    return name

def _activate_alzip_window(archive_name: str, timeout: float = 7.0) -> bool:
    """ALZip 창 활성화 및 최대화"""
    t0 = time.time()
    keywords = [archive_name, "ALZip", ".zip", ".alz"]
    while time.time() - t0 < timeout:
        wins = []
        try:
            for kw in keywords:
                wins.extend(gw.getWindowsWithTitle(kw))
        except Exception:
            pass

        uniq, seen = [], set()
        for w in wins:
            try:
                wid = (w._hWnd if hasattr(w, "_hWnd") else id(w))
                if wid in seen:
                    continue
                seen.add(wid)
                uniq.append(w)
            except Exception:
                continue

        for w in uniq:
            try:
                w.activate(); time.sleep(0.12)
                if hasattr(w, "isMinimized") and w.isMinimized:
                    w.restore(); time.sleep(0.15)
                # 최대화 (Alt+Space -> X)
                gui.hotkey("alt", "space"); time.sleep(0.10); gui.press("x"); time.sleep(0.15)
                return True
            except Exception:
                continue

        gui.hotkey("alt", "tab"); time.sleep(0.20)
        time.sleep(0.15)
    return False

def _center_of_active_window() -> Tuple[int, int]:
    """
    현재 활성 윈도우의 중심 좌표 반환.
    """
    try:
        w = gw.getActiveWindow()
        if w is None:
            screen_w, screen_h = gui.size()
            return (screen_w // 2, screen_h // 2)
        left, top = w.left, w.top
        width, height = w.width, w.height
        cx = left + width // 2
        cy = top + height // 2
        return (cx, cy)
    except Exception:
        screen_w, screen_h = gui.size()
        return (screen_w // 2, screen_h // 2)

def _press_many(key: str, n: int, delay: float = DEFAULT_PRESS_DELAY):
    for _ in range(max(0, n)):
        gui.press(key)
        time.sleep(delay)

# ---------------- 캡처 (활성 윈도우 기준 크롭) ----------------
def _screenshot_save(png_path: str,
                     top_frac_start: float = 1/5,
                     height_frac: float = 2/5):
    """
    활성 윈도우 기준으로 크롭하여 저장.
    top_frac_start: 창 높이의 시작 비율 (0.0 ~ 1.0)
    height_frac: 창 높이의 높이 비율 (0.0 ~ 1.0)
    """
    try:
        win = gw.getActiveWindow()
        if win and hasattr(win, "left"):
            left, top = win.left, win.top
            width, height = win.width, win.height

            crop_top = top + int(height * top_frac_start)
            crop_h = max(1, int(height * height_frac))

            screen_w, screen_h = gui.size()
            crop_left = max(0, left)
            crop_top = max(0, min(crop_top, screen_h - 1))
            crop_w = max(1, min(width, screen_w - crop_left))
            crop_h = max(1, min(crop_h, screen_h - crop_top))

            img = gui.screenshot(region=(crop_left, crop_top, crop_w, crop_h))
            img.save(png_path)
            return
    except Exception:
        pass

    img = gui.screenshot()
    img.save(png_path)

# ---------------- 공개 API ----------------
def capture_zip_with_alzip(
    zip_path: str,
    target_subdirs: List[str],
    out_base_dir: Optional[str] = None,
    open_wait: float = DEFAULT_OPEN_WAIT,
    nav_wait: float = DEFAULT_NAV_WAIT,
    press_delay: float = DEFAULT_PRESS_DELAY,
    down_count_stage1: int = 3,
    down_count_stage2: int = 2,
    down_count_stage3: int = 3,
    crop_top_frac: float = 1/5,
    crop_height_frac: float = 2/5,
    alzip_exe: Optional[str] = None,
) -> dict:
    """
    파라미터로 속도/동작을 조절할 수 있는 ALZip 캡처 함수.
    - down_count_stage*: 각 단계에서 누를 'down' 횟수 (기본 3,2,3)
    - press_delay: 방향키 반복 사이 딜레이 (초)
    - nav_wait: Enter/진입 후 대기 (초)
    - crop_*: 크롭 비율 (활성 창 기준)
    """
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"ZIP 경로 없음: {zip_path}")

    out_dir = _resolve_out_dir(zip_path, out_base_dir)
    base_name = os.path.splitext(os.path.basename(zip_path))[0]

    # 1) ALZip 실행
    if alzip_exe and os.path.exists(alzip_exe):
        subprocess.Popen([alzip_exe, zip_path])
    else:
        os.startfile(zip_path)
    time.sleep(open_wait)

    # 2) ALZip 창 활성화
    if not _activate_alzip_window(base_name, timeout=8.0):
        return {"output_dir": out_dir, "captures": [], "error": "ALZip 창 활성화 실패"}

    results = []
    try:
        # 안전: 활성 창 중심 클릭으로 포커스 확보
        cx, cy = _center_of_active_window()
        gui.click(cx, cy)
        time.sleep(0.15)

        # --- 3) ↓ stage1 + Enter ---
        _press_many("down", down_count_stage1, delay=press_delay)
        gui.press("enter")
        time.sleep(nav_wait)

        # --- 4) ↓ stage2 + Enter ---
        _press_many("down", down_count_stage2, delay=press_delay)
        gui.press("enter")
        time.sleep(nav_wait)

        # --- 5) 캡처 첫 번째 (크롭) ---
        first_leaf = (target_subdirs[0].replace("\\", "/").strip("/").split("/")[-1]) if target_subdirs else "capture1"
        first_png = os.path.join(out_dir, f"{_sanitize(first_leaf)}.png")
        _screenshot_save(first_png, top_frac_start=crop_top_frac, height_frac=crop_height_frac)
        results.append({"subdir": target_subdirs[0] if target_subdirs else "", "image_path": first_png, "status": "ok"})

        time.sleep(0.12)

        # --- 6) Enter (상위로 이동 또는 .. 선택) ---
        gui.press("enter")
        time.sleep(nav_wait)

        # --- 7) ↓ stage3 + Enter ---
        _press_many("down", down_count_stage3, delay=press_delay)
        gui.press("enter")
        time.sleep(nav_wait)

        # --- 8) 캡처 두 번째 (크롭) ---
        second_leaf = (target_subdirs[1].replace("\\", "/").strip("/").split("/")[-1]) if len(target_subdirs) > 1 else "capture2"
        second_png = os.path.join(out_dir, f"{_sanitize(second_leaf)}.png")
        _screenshot_save(second_png, top_frac_start=crop_top_frac, height_frac=crop_height_frac)
        results.append({"subdir": target_subdirs[1] if len(target_subdirs) > 1 else "", "image_path": second_png, "status": "ok"})

    except Exception as e:
        results.append({"subdir": "", "image_path": "", "status": f"fail: {e}"})
    finally:
        # ALZip 종료
        time.sleep(0.15)
        gui.hotkey("alt", "f4")

    return {"output_dir": out_dir, "captures": results}
