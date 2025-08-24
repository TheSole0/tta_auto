# main.py
from excel import load_exam_numbers
from alzip_zip_capture import capture_zip_with_alzip

import os
import pyautogui as gui
import time

# pyautogui 안전설정
gui.FAILSAFE = True        # 좌상단 모서리에 마우스 이동하면 즉시 예외 발생(긴급정지)
gui.PAUSE = 0.02           # pyautogui 호출 사이의 기본 대기 (짧게 설정 가능)

# ---------------- 사용자 설정 (필요에 따라 조정) ----------------
# 파일/디렉토리
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# ZIP_PATH = r"C:\Users\dlwls\auto\(25.08)ECM\zips\GS-C-24-0082 (주)삼우이머션.zip"

ZIP_PATH = r"C:\Users\dlwls\auto\(25.08)ECM\zips\GS-C-25-0002 (주)YH데이타베이스.zip" #----------------예외처리 코드 필요--------------

TARGET_SUBDIRS = [r"4.시험/나.설계", r"4.시험/다.수행"]

# ALZip 실행 파일(기본 연결이 ALZip이면 None)
ALZIP_EXE = None  # r"C:\Program Files\ESTsoft\ALZip\ALZip.exe"

# 속도/동작 파라미터 (여기를 조정해서 빠르게/안정적으로 튜닝)
OPEN_WAIT = 1.5        # ALZip 실행 후 최초 대기 (초). 성능 낮으면 늘리기.
NAV_WAIT = 0.45        # Enter 이후 안정 대기 (초). 핵심 안정값.
PRESS_DELAY = 0.04     # 방향키 연속 입력 사이 지연(초).
DOWN1 = 3              # 1단계: ↓ 몇 번
DOWN2 = 2              # 2단계: ↓ 몇 번
DOWN3 = 3              # 3단계: ↓ 몇 번

# 캡처 크롭 비율 (활성 창 기준)
CROP_TOP_FRAC = 0    # 창 높이의 위에서 시작 지점(예: 1/5)
CROP_HEIGHT_FRAC = 2/5 # 잘라낼 높이 비율(예: 2/5)

# 배치 처리 옵션 (엑셀로 여러 ZIP 처리 시 사용)
ENABLE_BATCH = False
EXCEL_PATH = r"C:\Users\dlwls\auto\(25.08)ECM\data.xlsx"
SLEEP_BETWEEN_ZIPS = 1.0  # ZIP 간 대기 (초)

# ---------------- 실행 ----------------
if __name__ == "__main__":
    # (선택) 엑셀에서 시험번호 읽기 (배치 처리에 사용)
    exam_numbers = []
    try:
        if ENABLE_BATCH:
            exam_numbers = load_exam_numbers(EXCEL_PATH, search_rows=50)
            # print(f"[INFO] 엑셀에서 {len(exam_numbers)}개 시험번호 읽음")
    except Exception as e:
        print("[WARN] 엑셀 읽기 실패:", e)

    def _run_single(zip_path: str):
        print(f"[RUN] 처리: {zip_path}")
        res = capture_zip_with_alzip(
            zip_path=zip_path,
            target_subdirs=TARGET_SUBDIRS,
            out_base_dir=BASE_DIR,
            alzip_exe=ALZIP_EXE,
            open_wait=OPEN_WAIT,
            nav_wait=NAV_WAIT,
            press_delay=PRESS_DELAY,
            down_count_stage1=DOWN1,
            down_count_stage2=DOWN2,
            down_count_stage3=DOWN3,
            crop_top_frac=CROP_TOP_FRAC,
            crop_height_frac=CROP_HEIGHT_FRAC,
        )
        print(f"[OUTPUT DIR] {res.get('output_dir')}")
        if "error" in res:
            print("[ERROR]", res["error"])
        for c in res.get("captures", []):
            print(f"- {c.get('subdir')} -> {c.get('image_path')} ({c.get('status')})")
        return res

    if not ENABLE_BATCH:
        # 단일 ZIP 처리
        _run_single(ZIP_PATH)
    else:
        # 배치: 엑셀의 시험번호로 zip파일 경로 조합 (사용자 규칙에 맞게 수정)
        base_folder = os.path.dirname(ZIP_PATH)
        for code in exam_numbers:
            # 예: 엑셀 값이 'GS-C-25-0002 (주)YH데이타베이스' 형식이라 가정
            zip_name = f"{code}.zip"
            zip_path = os.path.join(base_folder, zip_name)
            if not os.path.exists(zip_path):
                print(f"[SKIP] {zip_path} 없음")
                continue
            _run_single(zip_path)
            time.sleep(SLEEP_BETWEEN_ZIPS)
