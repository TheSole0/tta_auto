# ECM 자동화 프로젝트

## 📖 개요
이 프로젝트는 **GS 시험인증(소프트웨어 품질인증) 업무 자동화**를 위한 Python 기반 툴셋입니다.  
수작업으로 진행되던 ECM 관련 엑셀/문서 처리, ZIP 캡처, HWP/Word 보고서 작성 과정을 자동화하여  
업무 효율성과 정확성을 높이는 것을 목표로 합니다.

---

## 🚀 주요 기능
- **엑셀/CSV 데이터 처리**
  - 시험번호, 업체명, 시료명, 버전 등을 자동 파싱 및 전처리
- **ZIP 캡처 자동화 (`alzip_zip_capture.py`)**
  - 알집(Alzip) 프로그램을 통한 폴더 캡처 자동화
- **문서 자동화**
  - `doc.py`: Word 문서 자동 편집 및 데이터 치환
  - `hwp.py`: HWP 문서 자동 이미지 삽입
- **ECM 웹 자동화 (`ecm.py`)**
  - ECM 페이지 로그인 및 세션 키 관리 (`ecm_session_key.py`)
- **Notebook 예시 (`doc.ipynb`)**
  - 데이터 처리, 문서 삽입 기능 실험 및 시각화 예시 포함

---

## 📂 디렉터리 구조
```
(25.08.21)ECM/
 ├─ __pycache__/         # Python 캐시 (gitignore)
 ├─ GS-C-24-0082 ...     # 실제 데이터 (gitignore)
 ├─ GS-C-25-0002 ...     # 실제 데이터 (gitignore)
 ├─ zips/                # ZIP 캡처 산출물 (gitignore)
 ├─ 기록서/              # 시험 기록서 (gitignore)
 │   └─ 샘플/            # 템플릿/샘플
 ├─ alzip_zip_capture.py # 알집 ZIP 캡처 자동화
 ├─ doc.py               # Word 문서 자동화
 ├─ hwp.py               # HWP 자동화
 ├─ ecm.py               # ECM 웹 자동화
 ├─ ecm_session_key.py   # 세션 키 관리 (.env 필요)
 ├─ excel.py             # 엑셀 처리 모듈
 ├─ main.py              # 엔트리 포인트
 ├─ doc.ipynb            # 주피터 노트북 (예시/실험)
 └─ ...
```

---

## ⚙️ 설치 및 실행
### 1) 환경 세팅
```bash
# uv 사용 시
uv sync

# 또는 requirements.txt 기반
pip install -r requirements.txt
```

### 2) 환경 변수 파일 생성
루트에 `.env` 작성 (민감정보는 코드에 직접 넣지 않음):
```
ECM_USER_ID=your_id
ECM_PASSWORD=your_password
```

### 3) 실행
```bash
python main.py
```

---

## 📌 예시 워크플로우
1. `excel.py` → 시험번호/시료명/버전 데이터 로드
2. `alzip_zip_capture.py` → 관련 ZIP 폴더 자동 캡처
3. `doc.py`/`hwp.py` → 캡처 이미지 + 시험 데이터 자동 삽입
4. `ecm.py` → ECM 페이지 로그인 후 세션 확인

---

## ⚠️ 주의 사항
- `data.xlsx`, `*.csv`, `기록서/`, `zips/` 등 민감/대용량 파일은 GitHub에 올리지 않습니다.
- 원본 데이터 대신 **샘플 파일**을 `samples/` 폴더에 업로드하여 예시로 제공합니다.
- ECM 계정 정보는 `.env` 파일로 관리하며 코드 내에 직접 포함하지 않습니다.

---

## 📄 라이선스
이 프로젝트는 내부 업무 자동화 목적이며, 외부 공개 시 별도 라이선스를 지정해야 합니다.
