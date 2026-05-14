# 📊 대시보드 자동 갱신 도구

트래킹 시트(엑셀)에서 데이터를 읽어 **`dashboard.html`** 을 자동 생성합니다.

---

## ⚡ 30초 만에 시작하기

```cmd
cd C:\auto-tracker
python make_dashboard.py
```

→ `dashboard.html` 자동 생성됨. 더블클릭으로 열기!

---

## 📦 설치 (최초 1회만)

### 1. 파일 복사

이 폴더의 4개 파일을 **`C:\auto-tracker\` 안에** 넣으세요:

- `make_dashboard.py` ← 메인 실행 파일
- `extract_data.py`
- `render_html.py`
- `config.py`

### 2. 라이브러리 확인

이미 `update_tracker.py` 쓰고 계시면 **이미 설치되어 있어요** (openpyxl).
혹시 모르면:

```cmd
pip install openpyxl
```

### 3. config.py 확인

`config.py` 메모장으로 열어서 본인 환경에 맞는지 확인:

```python
# 엑셀 트래킹 시트 위치
EXCEL_PATH = r"C:\Users\jyu10\OneDrive\..."

# HTML 저장 위치
OUTPUT_PATH = r"C:\auto-tracker\dashboard.html"
```

→ 보통 그대로 두면 됩니다.

---

## 🚀 사용법

### 기본 사용

```cmd
cd C:\auto-tracker
python make_dashboard.py
```

→ 결과:
```
============================================================
📊 대시보드 자동 갱신
============================================================
📂 엑셀: C:\Users\jyu10\OneDrive\...
📄 출력: C:\auto-tracker\dashboard.html

⏳ 엑셀 데이터 읽는 중...
✅ 데이터 추출 완료
   현장 7개
   최근 활동 8건
   일별 추이 14일

📋 현장별 진행률:
   🟢 SIH2FC      651 /    670  ( 97.2%)  · 4개 항목
   🟡 CHW1FC      551 /   1213  ( 45.4%)  · 6개 항목
   🔴 ECH2FC        7 /    284  (  2.5%)  · 6개 항목
   ...

🎉 대시보드 생성 완료!
   📄 파일: C:\auto-tracker\dashboard.html
```

### 옵션

```cmd
# 특정 현장을 메인으로 (기본은 진행률 가장 높은 곳)
python make_dashboard.py --site CHW1FC

# 다른 위치에 저장
python make_dashboard.py --output C:\Users\Desktop\report.html

# 자세한 출력 안 함
python make_dashboard.py --quiet
```

---

## 🔄 매일 사용 흐름

### 옵션 1: 카톡 처리 후 수동 실행

```
1. 카톡 작업일보 받음
2. python update_tracker.py    ← 엑셀 갱신 (기존)
3. python make_dashboard.py    ← 대시보드 갱신 (NEW)
4. dashboard.html 더블클릭     ← 확인
```

### 옵션 2: 한 번에 자동 실행 (.bat)

`C:\auto-tracker\갱신.bat` 만들기:

```bat
@echo off
chcp 65001 > nul
cd /d "C:\auto-tracker"
echo [1/2] 트래킹 시트 갱신...
python update_tracker.py
echo.
echo [2/2] 대시보드 갱신...
python make_dashboard.py
echo.
pause
```

→ 더블클릭 한 번으로 둘 다 실행

---

## 🔮 나중에 Dropbox로 옮기기

지금: `C:\auto-tracker\dashboard.html` (본인 PC만)
   ↓
나중에: `C:\Users\jyu10\Dropbox\프로젝트 대시보드\dashboard.html` (자동 동기화)

`config.py`의 `OUTPUT_PATH` 한 줄만 바꾸면 됩니다:

```python
# 기존
OUTPUT_PATH = r"C:\auto-tracker\dashboard.html"

# 변경 후
OUTPUT_PATH = r"C:\Users\jyu10\Dropbox\프로젝트 대시보드\dashboard.html"
```

→ 이후 `python make_dashboard.py` 실행하면 Dropbox 폴더에 저장됨.

---

## 🐛 문제 해결

### "엑셀 파일을 찾을 수 없습니다"
→ `config.py`의 `EXCEL_PATH` 확인. 트래킹 시트 위치가 정확한지.

### "openpyxl 모듈 없음"
→ `pip install openpyxl` 실행

### HTML이 이상하게 보임
→ 인터넷이 켜져있는지 확인 (Pretendard 폰트, Chart.js 라이브러리 로딩)

### 데이터가 옛날 거임
→ 엑셀에서 변경 후 **저장**했는지 확인 (저장 안 하면 반영 X)

### 새 현장이 추가됐는데 안 보임
→ 새 시트가 엑셀에 추가되면 자동으로 잡힘. 다시 실행하면 OK.

---

## 📁 파일 구조

```
C:\auto-tracker\
├── update_tracker.py     ← 카톡 → 엑셀 (기존)
├── make_dashboard.py     ← 엑셀 → HTML (NEW) ⭐
├── extract_data.py       ← 데이터 추출 로직
├── render_html.py        ← HTML 렌더링 로직
├── config.py             ← 설정 (엑셀/출력 경로)
├── parser.py             ← 카톡 파서 (기존)
├── excel_updater.py      ← 엑셀 업데이트 (기존)
└── dashboard.html        ← 생성된 결과물 ✨
```

---

## 💡 자주 쓰는 명령어

```cmd
# 대시보드 새로 만들기
python make_dashboard.py

# CHW1FC를 메인으로 보기
python make_dashboard.py --site CHW1FC

# 결과만 보고 파일 생성 안 함 (테스트용 - 미지원, 그냥 실행)
python make_dashboard.py --quiet
```
