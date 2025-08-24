# data_file.py
import pandas as pd
from typing import List, Optional, Union


def _norm(s: str) -> str:
    if s is None:
        return ""
    return str(s).strip().replace("\n", "").replace("\r", "").replace(" ", "")

def load_exam_numbers(
    file_path: str,
    sheet_name: Optional[Union[str, int]] = 0,   # ✅ 수정  
    search_rows: int = 50,   # ✅ 탐색할 최대 행 수 (기본값 50)
) -> List[str]:
    """
    엑셀에서 '시험번호' 컬럼을 반환.
    - 헤더가 몇 행이든 자동 탐지
    - search_rows 매개변수로 탐색 범위 조절 가능
    """
    # 1) 전체를 헤더 없이 읽어, '시험번호'가 있는 행을 찾는다.
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str, engine="openpyxl")

    header_row_idx = None
    for i in range(min(search_rows, len(raw))):   # ✅ search_rows 사용
        row = raw.iloc[i].astype(str)
        if row.str.contains("시험번호", na=False).any():
            header_row_idx = i
            break

    if header_row_idx is None:
        preview = raw.head(3).to_dict(orient="list")
        raise ValueError(f"'시험번호' 헤더 행을 찾지 못했습니다. 미리보기: {preview}")

    # 2) 찾은 행을 헤더로 재지정해서 다시 읽기
    df = pd.read_excel(
        file_path, sheet_name=sheet_name, header=header_row_idx, dtype=str, engine="openpyxl"
    )

    # 3) 컬럼명 정규화
    orig_cols = list(df.columns)
    norm_cols = [_norm(c) for c in orig_cols]
    col_map = {n: o for n, o in zip(norm_cols, orig_cols)}

    # 4) '시험번호' 후보 키 찾기
    candidates = ["시험번호", "시험_번호", "시험-번호"]
    target = None
    for key in candidates:
        k = _norm(key)
        if k in col_map:
            target = col_map[k]
            break
    if target is None:
        for n, o in col_map.items():
            if "시험번호" in n:
                target = o
                break

    if target is None:
        raise ValueError(f"엑셀 파일에 '시험번호' 열이 존재하지 않습니다. 실제 열 이름: {orig_cols}")

    ser = df[target].astype(str).map(lambda x: x.strip())
    ser = ser[ser.notna() & (ser != "") & (ser != "nan")]
    return ser.tolist()
