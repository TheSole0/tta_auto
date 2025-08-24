# insert_images_to_hwp.py
import os
import time
from io import BytesIO
import ctypes

from PIL import Image
from win32com.client import gencache
import pywintypes

# ↓↓↓ 추가: 키보드/클립보드/포커스용
import pyautogui as gui
import win32clipboard
import win32con

gui.PAUSE = 0.05  # 키 입력 간 기본 지연

# --------- 사용자 기본값 ----------
HWP_PATH = r"C:\Users\dlwls\auto\(25.08)ECM\시험 기록서_GS-X-XX-0XXX(샘플_영남)_20250819.hwp"
IMAGES_FOLDER = r"C:\Users\dlwls\auto\(25.08)ECM\GS-C-25-0002 (주)YH데이타베이스"
OUT_SUFFIX = "_with_imgs"      # 저장파일명 접미사
MOVE_DOWN_COUNT = 22           # 아래 방향키 횟수 (각 이미지마다 이만큼 이동 후 붙여넣기)
PASTE_ONLY = True              # True: 아래로 MOVE_DOWN_COUNT회 + 붙여넣기(이미지 여러개 반복) 수행
# --------------------------------------------------------

# 지원 이미지 확장자
IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}

# 잠깐 대기(한/컴 반응속도 보정용)
INSERT_SLEEP = 0.12


def connect_hwp():
    try:
        hwp = gencache.EnsureDispatch("HwpFrame.HwpObject")
    except Exception as e:
        raise RuntimeError("한글 COM 연결 실패 (EnsureDispatch). 한글이 설치되어 있고 pywin32가 올바르게 설치됐는지 확인하세요.") from e

    try:
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    except Exception:
        pass

    try:
        hwp.XHwpWindows.Item(0).Visible = True
    except Exception:
        pass

    return hwp


def open_hwp(hwp, hwp_path):
    abspath = os.path.abspath(hwp_path)
    if not os.path.exists(abspath):
        raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {abspath}")

    try:
        hwp.Open(abspath)
    except pywintypes.com_error as ce:
        print("한글 Open에서 com_error 발생:", ce)
        print("-> gencache 캐시 재생성 후 재시도")
        try:
            gencache.EnsureDispatch("HwpFrame.HwpObject")
            time.sleep(0.2)
            hwp.Open(abspath)
        except Exception as e2:
            raise RuntimeError(f"HWP 열기 실패(두번째 시도 포함): {e2}") from ce


def save_hwp(hwp, original_path):
    base, ext = os.path.splitext(original_path)
    out_path = base + OUT_SUFFIX + ext
    try:
        hwp.SaveAs(os.path.abspath(out_path))
    except Exception:
        hwp.Save()
        out_path = original_path
    return out_path


def find_and_insert(hwp, placeholder, img_path):
    try:
        find_params = hwp.HParameterSet.HFindReplace
        try:
            find_params.SetItem("FindString", placeholder)
        except Exception:
            try:
                find_params.FindString = placeholder
            except Exception:
                pass
        hwp.HAction.Execute("Find", find_params)
    except Exception:
        raise RuntimeError(f"문서 검색(Find) 실패: 플레이스홀더 {placeholder}")

    try:
        ins = hwp.HParameterSet.HInsertPicture
        abs_img = os.path.abspath(img_path)
        try:
            ins.SetItem("FileName", abs_img)
            ins.SetItem("Width", 0)
            ins.SetItem("Height", 0)
        except Exception:
            try:
                ins.FileName = abs_img
                ins.Width = 0
                ins.Height = 0
            except Exception:
                pass
        hwp.HAction.Execute("InsertPicture", ins)
        time.sleep(INSERT_SLEEP)
        return True
    except Exception:
        try:
            getattr(hwp, "InsertPicture")(abs_img)  # type: ignore[attr-defined]
            time.sleep(INSERT_SLEEP)
            return True
        except Exception as e:
            raise RuntimeError(f"이미지 삽입 실패: {e}")


def insert_at_end(hwp, img_path):
    try:
        try:
            hwp.HAction.Run("MoveDocEnd")
        except Exception:
            pass
    except Exception:
        pass

    try:
        hwp.HAction.Run("InsertParagraph")
    except Exception:
        try:
            hwp.Run("InsertText", "\n")
        except Exception:
            pass

    abs_img = os.path.abspath(img_path)
    try:
        ins = hwp.HParameterSet.HInsertPicture
        try:
            ins.SetItem("FileName", abs_img)
            ins.SetItem("Width", 0)
            ins.SetItem("Height", 0)
        except Exception:
            try:
                ins.FileName = abs_img
                ins.Width = 0
                ins.Height = 0
            except Exception:
                pass
        hwp.HAction.Execute("InsertPicture", ins)
        time.sleep(INSERT_SLEEP)
        return True
    except Exception:
        try:
            getattr(hwp, "InsertPicture")(abs_img)  # type: ignore[attr-defined]
            time.sleep(INSERT_SLEEP)
            return True
        except Exception as e:
            raise RuntimeError(f"문서 끝 삽입 실패: {e}")


def collect_images_from_folder(folder):
    files = []
    for fname in os.listdir(folder):
        if os.path.splitext(fname)[1].lower() in IMG_EXTS:
            files.append(os.path.join(folder, fname))
    files.sort()
    return files


# ====== "아래로 N회 이동 → 붙여넣기" 유틸 ======
def _focus_hwp_window(hwp):
    try:
        hwnd = None
        try:
            hwnd = hwp.XHwpWindows.Item(0).HWnd
        except Exception:
            try:
                hwnd = hwp.XHwpWindows.Item(0).Hwnd
            except Exception:
                pass
        if hwnd:
            SW_SHOW = 5
            ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.15)
    except Exception:
        pass


def _set_image_to_clipboard(image_path: str):
    img = Image.open(image_path).convert("RGB")
    output = BytesIO()
    img.save(output, format="BMP")
    data = output.getvalue()[14:]  # BMP 헤더 제거 -> DIB
    output.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


def _move_down_and_paste(hwp, count: int = 22):
    # COM 시도: 아래로 이동 후 Paste
    try:
        for _ in range(count):
            try:
                hwp.HAction.Run("MoveDown")
            except Exception:
                try:
                    hwp.HAction.Run("MoveLineDown")
                except Exception:
                    pass
            time.sleep(0.02)
        try:
            hwp.HAction.Run("Paste")
            return True
        except Exception:
            pass
    except Exception:
        pass

    # fallback: 키 이벤트로 수행
    _focus_hwp_window(hwp)
    gui.press('down', presses=count, interval=0.02)
    gui.hotkey('ctrl', 'v')
    return True
# ====== 유틸 끝 ======


def main(hwp_path, images_folder):
    hwp = connect_hwp()
    open_hwp(hwp, hwp_path)

    images = collect_images_from_folder(images_folder)
    if not images:
        print("이미지 폴더에 지원되는 이미지 파일이 없습니다.")
        try:
            hwp.Quit()
        except Exception:
            pass
        return

    # ---- 수정된 PASTE_ONLY 동작: 모든 이미지에 대해 순차 삽입 수행 ----
    if PASTE_ONLY:
        inserted = []
        failed = {}
        for idx, img in enumerate(images, start=1):
            try:
                _set_image_to_clipboard(img)
                # 클립보드가 안정화될 수 있도록 잠시 대기
                time.sleep(0.12)
                success = _move_down_and_paste(hwp, count=MOVE_DOWN_COUNT)
                if success:
                    inserted.append(img)
                    print(f"[OK] 삽입: {img}")
                else:
                    failed[img] = "붙여넣기 실패(unknown)"
                    print(f"[WARN] 붙여넣기 실패: {img}")
                # 이미지 간 간격(필요시 늘리세요)
                time.sleep(0.2)
            except Exception as e:
                failed[img] = str(e)
                print(f"[ERR] 처리 실패({img}): {e}")

        out_path = save_hwp(hwp, hwp_path)
        print("저장 완료:", out_path)

        try:
            hwp.Quit()
        except Exception:
            pass

        # 요약
        print("요약:")
        print(f" - 총 이미지 파일: {len(images)}")
        print(f" - 정상 삽입: {len(inserted)}")
        print(f" - 실패: {len(failed)}")
        if failed:
            for k, v in failed.items():
                print("   *", k, ":", v)
        return

    # ---- 기존 placeholder 루프 (변경 없음) ----
    results = {"inserted_by_placeholder": [], "inserted_at_end": [], "failed": {}}

    for idx, img in enumerate(images, start=1):
        placeholder = f"{{IMG_{idx}}}"  # {IMG_1}, {IMG_2}, ...
        try:
            try:
                success = find_and_insert(hwp, placeholder, img)
                if success:
                    results["inserted_by_placeholder"].append((placeholder, img))
                    print(f"[OK] 삽입(플레이스홀더): {placeholder} <- {img}")
                    continue
            except Exception as e:
                print(f"[WARN] 플레이스홀더 삽입 실패({placeholder}): {e}")

            try:
                success = insert_at_end(hwp, img)
                if success:
                    results["inserted_at_end"].append(img)
                    print(f"[OK] 문서 끝에 삽입: {img}")
            except Exception as e2:
                print(f"[ERR] 문서 끝 삽입 실패: {e2}")
                results["failed"][img] = str(e2)

        except Exception as e_outer:
            results["failed"][img] = str(e_outer)
            print(f"[ERR] 처리 실패({img}): {e_outer}")

    out_path = save_hwp(hwp, hwp_path)
    print("저장 완료:", out_path)

    try:
        hwp.Quit()
    except Exception:
        pass

    print("요약:")
    print(" - 플레이스홀더로 삽입:", len(results["inserted_by_placeholder"]))
    print(" - 문서 끝에 삽입:", len(results["inserted_at_end"]))
    print(" - 실패:", len(results["failed"]))
    if results["failed"]:
        for k, v in results["failed"].items():
            print("   *", k, ":", v)


if __name__ == "__main__":
    main(HWP_PATH, IMAGES_FOLDER)
