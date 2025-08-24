from pathlib import Path
import pandas as pd

EXCEL_PATH = Path(r"C:\Users\dlwls\auto\(25.08)ECM\data.xlsx")

def load_excel_to_df(path: Path) -> pd.DataFrame:
    """
    단일 시트 엑셀을 전체 읽어 DataFrame 반환.
    - sheet_name=0 으로 첫 번째(유일한) 시트 읽음.
    - openpyxl 엔진 사용.
    """
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
    # header, index 등 기본 동작 그대로 사용 (단순하게)
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    return df

if __name__ == "__main__":
    df = load_excel_to_df(EXCEL_PATH)
    print("파일 경로:", EXCEL_PATH)
    print("데이터프레임 크기 (rows, cols):", df.shape)
    print("컬럼 목록:", list(df.columns))
    print("\n첫 5행 샘플:")
    print(df.head())
    # 필요하면 아래 주석 해제하여 CSV 또는 parquet로 저장 가능
    # df.to_csv(EXCEL_PATH.with_suffix(".csv"), index=False)
    # df.to_parquet(EXCEL_PATH.with_suffix(".parquet"), index=False)
