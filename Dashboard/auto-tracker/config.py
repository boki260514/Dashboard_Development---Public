"""
설정 파일 - 카톡 메시지와 엑셀 시트 사이의 매핑 정보를 관리합니다.

이 파일만 수정하면 됩니다:
  - 새 현장이 추가되면 SITE_MAPPING에 추가
  - 새 작업 종류가 추가되면 ROW_MAPPING에 추가
  - 엑셀 파일 경로가 바뀌면 EXCEL_PATH 수정
  - 대시보드 출력 위치가 바뀌면 OUTPUT_PATH 수정
"""

# ============================================================
# 엑셀 파일 경로 (트래킹 시트)
# ============================================================
import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))
EXCEL_PATH = _os.path.join(
    _os.path.dirname(_HERE),
    "HCNSSERVICE_Tracking_v2.xlsx",
)

# 백업 폴더 (엑셀 파일과 같은 폴더 안에 backups/ 자동 생성)
BACKUP_FOLDER_NAME = "backups"


# ============================================================
# 카톡 현장명 → 엑셀 시트명 매핑
# ============================================================
# 카톡에 다양한 표기로 올 수 있으니 여러 변형을 모두 등록해 둡니다.
# 키는 모두 소문자, 공백 제거된 형태로 저장됨 (자동 정규화).
SITE_MAPPING = {
    # 쿠팡 창원 1센터
    "쿠팡창원1센터": "CHW1FC",
    "창원1센터": "CHW1FC",
    "창원1": "CHW1FC",
    "chw1fc": "CHW1FC",
    "chw1": "CHW1FC",

    # 쿠팡 창원 2센터
    "쿠팡창원2센터": "CHW2FC",
    "창원2센터": "CHW2FC",
    "창원2": "CHW2FC",
    "chw2fc": "CHW2FC",
    "chw2": "CHW2FC",

    # 쿠팡 시흥 2센터
    "쿠팡시흥2센터": "SIH2FC",
    "시흥2센터": "SIH2FC",
    "시흥2": "SIH2FC",
    "sih2fc": "SIH2FC",
    "sih2": "SIH2FC",

    # 쿠팡 이천 2센터
    "쿠팡이천2센터": "ECH2FC",
    "이천2센터": "ECH2FC",
    "이천2": "ECH2FC",
    "ech2fc": "ECH2FC",
    "ech2": "ECH2FC",

    # 쿠팡 동탄 1센터 (추정 - 실제 매핑 확인 필요)
    "쿠팡동탄1센터": "DON1FC",
    "동탄1센터": "DON1FC",
    "동탄1": "DON1FC",
    "don1fc": "DON1FC",
    "don1": "DON1FC",

    # 쿠팡 마산 1센터 (추정 - 실제 매핑 확인 필요)
    "쿠팡마산1센터": "MCN1FC",
    "마산1센터": "MCN1FC",
    "마산1": "MCN1FC",
    "mcn1fc": "MCN1FC",
    "mcn1": "MCN1FC",
}


# ============================================================
# 카톡 작업명 → 엑셀 행 매핑 (각 시트별)
# ============================================================
# 카톡 메시지에 등장하는 라인/작업 키워드를
# 엑셀의 정확한 행 헤더 이름과 매칭합니다.
#
# 형식: 시트명 → { 카톡 키워드: 엑셀 행 헤더 }
# 카톡 키워드는 정규화(소문자, 공백 제거) 후 매칭됩니다.
ROW_MAPPING = {
    "CHW1FC": {
        # PB#01 Line - Gap Cover → PB#01 Gap guarding Installation
        "pb#1line_gapcover": "PB#01 Gap guarding Installation",
        "pb#01line_gapcover": "PB#01 Gap guarding Installation",
        "pb1line_gapcover": "PB#01 Gap guarding Installation",
        "pb#1_gapcover": "PB#01 Gap guarding Installation",
        "pb#01_gapcover": "PB#01 Gap guarding Installation",
        # PB#02 Line - Gap Cover → PB#02 Gap guarding Installation
        "pb#2line_gapcover": "PB#02 Gap guarding Installation",
        "pb#02line_gapcover": "PB#02 Gap guarding Installation",
        "pb2line_gapcover": "PB#02 Gap guarding Installation",
        "pb#2_gapcover": "PB#02 Gap guarding Installation",
        "pb#02_gapcover": "PB#02 Gap guarding Installation",
        # CB Line - Gap Cover → CB Gap guarding Installation
        "cbline_gapcover": "CB Gap guarding Installation",
        "cb_gapcover": "CB Gap guarding Installation",
        # PB#01 Line - Cover → PB#01 Cover Installation
        "pb#1line_cover": "PB#01 Cover Installation",
        "pb#01line_cover": "PB#01 Cover Installation",
        "pb#1_cover": "PB#01 Cover Installation",
        "pb#01_cover": "PB#01 Cover Installation",
        # PB#02 Line - Cover
        "pb#2line_cover": "PB#02 Cover Installation",
        "pb#02line_cover": "PB#02 Cover Installation",
        "pb#2_cover": "PB#02 Cover Installation",
        "pb#02_cover": "PB#02 Cover Installation",
        # CB Line - Cover
        "cbline_cover": "CB Cover Installation",
        "cb_cover": "CB Cover Installation",
    },

    "ECH2FC": {
        # Gap Cover (Gap guarding)
        "pb#1line_gapcover": "PB#01 Gap guarding Installation",
        "pb#01line_gapcover": "PB#01 Gap guarding Installation",
        "pb#1_gapcover": "PB#01 Gap guarding Installation",
        "pb#01_gapcover": "PB#01 Gap guarding Installation",
        "pb#2line_gapcover": "PB#02 Gap guarding Installation",
        "pb#02line_gapcover": "PB#02 Gap guarding Installation",
        "pb#2_gapcover": "PB#02 Gap guarding Installation",
        "pb#02_gapcover": "PB#02 Gap guarding Installation",
        "cbline_gapcover": "CB Gap guarding Installation",
        "cb_gapcover": "CB Gap guarding Installation",
        # Cable (Cable Management Parts)
        "pb#1line_cable": "PB#01 Cable Installation",
        "pb#01line_cable": "PB#01 Cable Installation",
        "pb#1line_하네스케이블": "PB#01 Cable Installation",
        "pb#1_cable": "PB#01 Cable Installation",
        "pb#01_cable": "PB#01 Cable Installation",
        "pb#01_하네스케이블": "PB#01 Cable Installation",
        "pb#1_하네스케이블": "PB#01 Cable Installation",
        "pb#2line_cable": "PB#02 Cable Installation",
        "pb#02line_cable": "PB#02 Cable Installation",
        "pb#2line_하네스케이블": "PB#02 Cable Installation",
        "pb#2_cable": "PB#02 Cable Installation",
        "pb#02_cable": "PB#02 Cable Installation",
        "cbline_cable": "CB Cable Installation",
        "cb_cable": "CB Cable Installation",
        "cb_하네스케이블": "CB Cable Installation",
        "cbline_하네스케이블": "CB Cable Installation",
    },

    "SIH2FC": {
        # PB Gap guarding (PB#01/PB#02 구분 없이 'PB'만)
        "pbline_gapcover": "PB Gap guarding Installation",
        "pb_gapcover": "PB Gap guarding Installation",
        "pb#1_gapcover": "PB Gap guarding Installation",
        "pb#01_gapcover": "PB Gap guarding Installation",
        "pb#2_gapcover": "PB Gap guarding Installation",
        "pb#02_gapcover": "PB Gap guarding Installation",
        # CB Gap guarding
        "cbline_gapcover": "CB Gap guarding Installation",
        "cb_gapcover": "CB Gap guarding Installation",
        # CB Cover
        "cbline_cover": "CB Cover Installation",
        "cb_cover": "CB Cover Installation",
        # CB Gap Plate
        "cbline_gapplate": "CB Gap Plate Installation",
        "cb_gapplate": "CB Gap Plate Installation",
        "cb_플레이트": "CB Gap Plate Installation",
        "cbline_플레이트": "CB Gap Plate Installation",
    },

    "CHW2FC": {
        # CB Cable Installation만 있음
        "cbline_cable": "CB Cable Installation",
        "cb_cable": "CB Cable Installation",
        "cb_하네스케이블": "CB Cable Installation",
        "cbline_하네스케이블": "CB Cable Installation",
    },

    "DON1FC": {
        # ST#01~04 Cable Installation
        "st#1line_cable": "ST#01 Cable Installation",
        "st#01line_cable": "ST#01 Cable Installation",
        "st1line_cable": "ST#01 Cable Installation",
        "st#1_cable": "ST#01 Cable Installation",
        "st#01_cable": "ST#01 Cable Installation",
        "st#1_하네스케이블": "ST#01 Cable Installation",
        "st#01_하네스케이블": "ST#01 Cable Installation",

        "st#2line_cable": "ST#02 Cable Installation",
        "st#02line_cable": "ST#02 Cable Installation",
        "st#2_cable": "ST#02 Cable Installation",
        "st#02_cable": "ST#02 Cable Installation",
        "st#2_하네스케이블": "ST#02 Cable Installation",
        "st#02_하네스케이블": "ST#02 Cable Installation",

        "st#3line_cable": "ST#03 Cable Installation",
        "st#03line_cable": "ST#03 Cable Installation",
        "st#3_cable": "ST#03 Cable Installation",
        "st#03_cable": "ST#03 Cable Installation",
        "st#3_하네스케이블": "ST#03 Cable Installation",
        "st#03_하네스케이블": "ST#03 Cable Installation",

        "st#4line_cable": "ST#04 Cable Installation",
        "st#04line_cable": "ST#04 Cable Installation",
        "st#4_cable": "ST#04 Cable Installation",
        "st#04_cable": "ST#04 Cable Installation",
        "st#4_하네스케이블": "ST#04 Cable Installation",
        "st#04_하네스케이블": "ST#04 Cable Installation",

        # PB#01 Cable Installation
        "pb#1line_cable": "PB#01 Cable Installation",
        "pb#01line_cable": "PB#01 Cable Installation",
        "pb#1_cable": "PB#01 Cable Installation",
        "pb#01_cable": "PB#01 Cable Installation",
        "pb#1_하네스케이블": "PB#01 Cable Installation",
        "pb#01_하네스케이블": "PB#01 Cable Installation",
    },

    "MCN1FC": {
        # CB#01, CB#02 Cable Installation
        "cb#1line_cable": "CB#01 Cable Installation",
        "cb#01line_cable": "CB#01 Cable Installation",
        "cb#1_cable": "CB#01 Cable Installation",
        "cb#01_cable": "CB#01 Cable Installation",
        "cb#1_하네스케이블": "CB#01 Cable Installation",
        "cb#01_하네스케이블": "CB#01 Cable Installation",

        "cb#2line_cable": "CB#02 Cable Installation",
        "cb#02line_cable": "CB#02 Cable Installation",
        "cb#2_cable": "CB#02 Cable Installation",
        "cb#02_cable": "CB#02 Cable Installation",
        "cb#2_하네스케이블": "CB#02 Cable Installation",
        "cb#02_하네스케이블": "CB#02 Cable Installation",
    },
}


# ============================================================
# 작업 종류 키워드 (라인 + 작업명 조합용)
# ============================================================
# 카톡 메시지에서 작업 종류를 식별하는 키워드들
WORK_TYPE_KEYWORDS = {
    "gapcover": ["gap cover", "gapcover", "갭커버", "갭 커버"],
    "cover": ["cover", "커버"],  # gap cover와 구분 필요 - gapcover 먼저 체크
    "cable": [
        "하네스 케이블", "하네스케이블", "케이블",
        "cable", "harness cable", "harness",
    ],
    "gapplate": ["gap plate", "gapplate", "갭 플레이트", "플레이트"],
}


# ============================================================
# 엑셀 행 구조 (모든 시트 공통)
# ============================================================
# 각 작업 항목 그룹은 5개 행으로 구성:
#   행 N+0: 항목명 (Plan)
#   행 N+1: Actual           ← 수량 입력 ⭐
#   행 N+2: Downtime window hours  ← 다운타임 입력
#   행 N+3: Installation / Hour    ← 시간당 설치량 (자동 계산)
#   행 N+4: Labor                  ← 투입 인원 입력
ROW_OFFSET = {
    "actual": 1,            # 작업 수량
    "downtime": 2,          # 다운타임 (시간)
    "installation_hour": 3, # 시간당 설치량 (자동 - 보통 안 채움)
    "labor": 4,             # 투입 인원
}


# ============================================================
# 날짜 행 위치 (모든 시트 공통)
# ============================================================
MONTH_ROW = 7   # 월 헤더 (April / May / June / ...)
DAY_ROW = 8     # 일자 (1, 2, 3, ..., 30, 31)
DATA_START_COL = 4  # D열부터 데이터 시작


# ============================================================
# 대시보드 생성기 설정 (make_dashboard.py에서 사용)
# ============================================================
# 1단계 (지금): 본인 PC에 저장
OUTPUT_PATH = _os.path.join(_HERE, "dashboard.html")

# 2단계 (나중): Dropbox 폴더로 변경하면 자동 동기화
# OUTPUT_PATH = r"C:\Users\jyu10\Dropbox\프로젝트 대시보드\dashboard.html"

# 메인에 표시할 현장 (None이면 누적 작업이 가장 많은 현장 자동 선택)
DEFAULT_ACTIVE_SITE = None


# ============================================================
# 현장별 총 계획 멘데이(M-day)
# ============================================================
# 엑셀 트래킹 시트의 Labor 행에는 실제 사용 인원만 기록되고
# '총 계획 멘데이'는 별도로 저장되지 않으므로 여기서 관리합니다.
#
# 값을 설정하지 않거나 0이면 대시보드에 '-'로 표시됩니다.
# 값을 채우면 진행률(%) 옆에 '총 계획 / 누적 사용 / 금일 사용' 으로 표시됩니다.
#
# 예) "CHW1FC": 150,   # 창원 1센터 총 계획 150 멘데이
PLAN_MDAY_BY_SITE = {
    "CHW1FC": 0,
    "CHW2FC": 0,
    "SIH2FC": 0,
    "ECH2FC": 0,
    "DON1FC": 0,
    "MCN1FC": 0,
    "None PJT": 0,
}
