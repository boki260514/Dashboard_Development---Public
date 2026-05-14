# CLAUDE.md — auto-tracker 시스템 컨텍스트

> **목적:** 다른 사람/세션이 이 디렉터리에서 작업을 이어받을 때 빠르게 파악할 수 있도록
> 시스템 구조·데이터 규칙·환경 설정·알려진 이슈를 누적 기록한다.
> 사용자 매뉴얼이 아니라 **AI/협업용 컨텍스트 문서**다.
> 코드를 바꿀 때마다 이 문서도 함께 갱신할 것.

---

## 1. 디렉터리 위치

작업 루트: `C:\Users\jyu10\한국씨엔에스(주)\Dashboard_Development - Dashboard\auto-tracker\`

이 디렉터리에는 **두 개의 독립 파이프라인**이 동거한다.

| 파이프라인 | 메인 스크립트 | 처리 대상 | 출력 |
|---|---|---|---|
| ① 작업일보 자동화 (구) | `email_fetcher.py` | 특정 발신자 + '업무일지' 제목 메일 | `../HCNSSERVICE_Tracking_v2.xlsx` |
| ② 일반 메일 로그 (신) | `mail_log_fetcher.py` | INBOX 전체 업무 메일 | `../Mail_Log.xlsx` + 작업일보 첨부도 ①과 동일 트래커에 함께 반영 |

두 파이프라인 모두 같은 IMAP 계정·같은 `.env`·같은 트래커 파일을 공유한다.

---

## 2. 시스템 흐름

```
                ┌──────────────────────────────────────────────┐
                │ 다우오피스 IMAP (imap.daouoffice.com:993)    │
                └──────────────────────────────────────────────┘
                                  │
       ┌──────────────────────────┴──────────────────────────┐
       ↓                                                     ↓
┌──────────────┐                                  ┌──────────────────┐
│ ① email_fetcher (작업일보 전용)                  │ ② mail_log_fetcher (일반 업무 메일)
│  - 발신자/제목 필터                                │   - 광고/시스템 1차 필터
│  - excel_parser → DailyReport                    │   - Claude Sonnet 4.6 분류 (20필드)
│  - excel_updater.update_excel                    │   - 첨부 .xlsx도 ①과 동일 흐름
│  - 미수신 시 SMTP 알림                            │   - Mail_Log 시트 append
└──────────────┘                                  └──────────────────┘
       │                                                     │
       ↓                                                     ↓
┌──────────────────────────────┐              ┌──────────────────────────────┐
│ HCNSSERVICE_Tracking_v2.xlsx │              │ Mail_Log.xlsx                │
│  · 현장별 시트 (CHW1FC 등)    │              │  · Mail_Log 시트 (20컬럼)     │
│  · backups/ 자동 백업          │              │  · Run_Log 시트              │
└──────────────────────────────┘              └──────────────────────────────┘
       │                                                     │
       ↓                                                     ↓
┌──────────────────────────────┐              ┌──────────────────────────────┐
│ make_dashboard.py → HTML 7종 │              │ mail_dashboard.py → HTML 1종 │
│  dashboard.html              │              │  mail_dashboard.html         │
│  dashboard_<SITE>.html       │              │                              │
└──────────────────────────────┘              └──────────────────────────────┘
```

---

## 3. 파일 역할 (한 줄)

### 파이프라인 ①: 작업일보 (기존)
- `email_fetcher.py` — IMAP 검색·첨부 다운로드·미수신 알림 (SMTP).
- `excel_parser.py` — 작업일보 첨부(.xlsx) → `DailyReport` 변환. **C2=공사명, C3=일자, B7=수량텍스트, E38~E42=인원** 가정.
- `excel_updater.py` — `DailyReport` → `HCNSSERVICE_Tracking_v2.xlsx` 셀 단위 업데이트 + `backups/` 자동.
- `parser.py` — 카톡 메시지 → `DailyReport` 변환 (구 카톡 경로). `update_tracker.py`와 함께 사용.
- `update_tracker.py` — 카톡 페이스트 → 트래커 업데이트 CLI.
- `make_dashboard.py` — 트래커 → `dashboard.html` + 현장별 HTML.
- `extract_data.py` — 트래커에서 시트별 데이터 추출 (대시보드용).
- `render_html.py` — 대시보드 HTML 렌더링.
- `watch_excel.py` — 트래커 파일 변경 감시 → 자동 대시보드 재생성.
- `config.py` — 트래커 경로, 현장 매핑, 행/열 오프셋, 작업종류 키워드 사전.

### 파이프라인 ②: 일반 메일 로그 (신규)
- `mail_log_fetcher.py` — **메인.** IMAP → 광고필터 → Claude 분류 → Mail_Log append → 대시보드 재생성.
  - 첨부 .xlsx가 작업일보 형식이면 ①의 `excel_parser`+`excel_updater` 재사용해 트래커에도 반영.
- `mail_classifier.py` — Claude Sonnet 4.6 호출, JSON 추출/정규화, 광고 1차 필터, fallback(검토필요).
- `mail_log_excel.py` — Mail_Log.xlsx 읽기/쓰기, 중복키 로드, 파일잠금 fallback(`*_locked_*.xlsx`).
- `mail_log_schema.py` — Mail_Log 20컬럼·Run_Log 11컬럼 정의 + 빈 템플릿 생성기.
- `mail_dashboard.py` — Mail_Log → `mail_dashboard.html` (7 지표).

### 수동 첨부 일괄 처리 도구 (대용량 첨부 등)
- `batch_import_attachments.py` — `오늘_첨부/` 폴더의 .xlsm/.xlsx를 날짜순으로 일괄 트래커 반영.
  - 다우오피스 대용량 첨부(.zip 링크) 등 IMAP 자동 수신 불가 케이스에 사용.
  - 첫 파일만 자동 백업 (한 세션 = 한 백업), 나머지는 skip.
  - 사용: `python batch_import_attachments.py [--dry-run] [--dir 폴더]`
- `merge_tracker_from_email.py` — 이메일로 통째 트래커 사본을 받았을 때 안전 머지 도구.
  - 머지 규칙: 첨부=값/트래커=빈 셀만 추가, 충돌은 자동 변경 없이 리스트로만 보고.
  - 현재 보류 상태(레이아웃 차이 발견됨 — §7 참조).

### 공통 / 운영
- `.env` (gitignore) — IMAP/SMTP 자격, 허용 발신자, 제목 키워드, `ANTHROPIC_API_KEY`.
- `.env.example` — `.env` 템플릿.
- `requirements.txt` — `openpyxl`, `pyperclip`, `anthropic`.
- `mail_state.json` — `last_success_ts`, `run_seq`, `last_run_id`. 첫 실행 시 자동 생성.

### 배치 파일 (사용자가 더블클릭)
- `자동화_등록.bat` — 작업일보 스케줄 등록. **매일 08:00 시작, 30분 간격, 2시간 (8:00·8:30·9:00·9:30·10:00)**.
  - 작업명: `HCNS_메일일보처리_매일오전`
- `자동화_해제.bat` — 위 작업 + 옛 이름(`*_매일오전9시`) 모두 정리.
- `자동화_등록_메일로그.bat` — 메일로그 스케줄 등록. **매일 08:05 시작, 30분 간격, 2시간 (8:05·8:35·9:05·9:35·10:05)**.
  - 작업명: `HCNS_메일로그처리_매일오전`
- `자동화_해제_메일로그.bat` — 위 + 옛 이름 정리.
- `메일일보처리.bat` — 파이프라인 ① 1회 실행 래퍼.
- `메일로그처리.bat` — 파이프라인 ② 1회 실행 래퍼.
- `대시보드_갱신.bat` / `대시보드_감시시작.bat` / `대시보드_감시중지.bat` — 대시보드 갱신/감시.

---

## 4. 데이터 규칙

### 4.1 HCNSSERVICE_Tracking_v2.xlsx 구조 (config.py 기준)
- 시트명 = 현장 코드 (`CHW1FC`, `CHW2FC`, `SIH2FC`, `ECH2FC`, `DON1FC`, `MCN1FC`).
- 7행 = 월 헤더 (`April`/`May`/...). 8행 = 일자(1~31). 4열(D)부터 데이터 시작.
- 각 작업 항목 그룹은 **5개 행 구성**:
  - +0: 항목명(Plan)
  - +1: **Actual** ← 수량 ⭐
  - +2: Downtime window hours
  - +3: Installation / Hour (자동 계산용 — 보통 안 채움)
  - +4: **Labor** ← 투입 인원

### 4.2 ROW_MAPPING (config.py)
카톡/메일 키워드 → 엑셀 행 헤더 매핑. 현장별로 분리.
키는 모두 **소문자·공백제거 정규화** 후 매칭됨.

예 — CHW1FC:
- `pb#1line_gapcover` → `PB#01 Gap guarding Installation`
- `cbline_cover` → `CB Cover Installation`

### 4.3 SITE_MAPPING (config.py)
카톡/메일 본문의 현장 표기 → 시트명. 한글/영문 변형 다수 등록 (예: "창원1센터", "쿠팡창원1센터", "chw1fc" 등 → `CHW1FC`).

### 4.4 WORK_TYPE_KEYWORDS
- `gapcover` — "gap cover", "갭커버", ...
- `cover` — "cover", "커버" (단 gapcover 먼저 매칭 필요)
- `cable` — "케이블", "하네스케이블", "harness", "harness cable"
- `gapplate` — "gap plate", "갭 플레이트", "플레이트"

### 4.5 Mail_Log.xlsx 컬럼 (mail_log_schema.py)
순서가 컬럼 위치. 헤더는 1행, 데이터는 2행부터.

| # | 컬럼 | 형식 |
|---|---|---|
| 1 | 메일ID | `Message-ID` 우선, 없으면 `UID-xxx` |
| 2 | 수신일시 | `YYYY-MM-DD HH:mm:ss` |
| 3-4 | 발신자이름/이메일 | text |
| 5 | 회사명 | classifier 추정 |
| 6 | 제목 | 원문 |
| 7 | 업무유형 | 견적/주문/계약/정산/CS/일정/보고/기타 |
| 8 | 핵심내용요약 | 1~3문장 |
| 9 | 요청사항 | text |
| 10 | 금액 | 숫자(원). 빈 값 `""` |
| 11 | 수량 | 숫자 |
| 12 | 납기일 | `YYYY-MM-DD` |
| 13 | 담당자 | text |
| 14 | 우선순위 | 긴급/높음/보통/낮음 |
| 15 | 상태 | 신규/진행중/검토필요/완료 |
| 16 | 첨부파일명 | 콤마 구분 |
| 17 | 원문참조 | IMAP UID |
| 18 | 처리일시 | `YYYY-MM-DD HH:mm:ss` |
| 19 | 신뢰도 | 높음/중간/낮음 |
| 20 | 비고 | "트래커 반영: SITE date N셀" 등 메타 정보 누적 |

### 4.6 중복 판정 우선순위 (mail_log_excel.make_fingerprint)
1. 메일ID (Message-ID 또는 UID-xxx)
2. fallback: `수신일시 | 발신자이메일 | 제목`

### 4.7 사람이 직접 수정한 행은 보존
중복 판정이 메일ID 기준이라, 사람이 회사명/담당자/상태/비고 등을 수정해도 다음 실행 때 덮어쓰지 않는다.

---

## 5. 환경 설정 (.env 키)

| 키 | 용도 | 비고 |
|---|---|---|
| `EMAIL_ADDRESS` | IMAP·SMTP 로그인 ID | 다우오피스 메일 |
| `EMAIL_PASSWORD` | 패스워드 | **로그·출력 절대 금지** |
| `IMAP_HOST` / `IMAP_PORT` | 메일 수신 | 기본 `imap.daouoffice.com:993` |
| `SMTP_HOST` / `SMTP_PORT` | 미수신 알림 발송 | 기본 `smtp.daouoffice.com:465` |
| `ALLOWED_SENDERS` | 작업일보 발신자 화이트리스트 (콤마 구분) | 파이프라인 ①에만 사용 |
| `SUBJECT_KEYWORD` | 작업일보 제목 키워드 | **현재 값: `작업일보`** (2026-05-14 수정. 이전 `업무일지`는 일일 업무 리스트라 수량 없음) |
| `ALERT_TO` | 미수신 알림 받을 주소 | 보통 본인 |
| `ANTHROPIC_API_KEY` | Claude API | 비우면 파이프라인 ②가 모든 메일을 `검토필요`로 기록 |

`.env`는 `.gitignore` 처리, 깃에 절대 커밋 금지.

---

## 6. 스케줄러 작업명 (Windows Task Scheduler)

| 작업명 | 트리거 | 실행 BAT |
|---|---|---|
| `HCNS_메일일보처리_매일오전` | 매일 08:00, 30분 간격, 2시간 동안 (총 5회) | `메일일보처리.bat` |
| `HCNS_메일로그처리_매일오전` | 매일 08:05, 30분 간격, 2시간 동안 (총 5회) | `메일로그처리.bat` |

옛 작업명(`*_매일오전9시`)이 남아있다면 해제 BAT이 함께 정리한다.

확인: `Win+R → taskschd.msc → 작업 스케줄러 라이브러리`

---

## 7. 알려진 이슈 / 호환성 메모

### 7.1 작업일보 양식 vs 일일 업무일지 양식 (해결됨 2026-05-14)
- "작업일보" ≠ "업무일지". 둘은 별개 메일이고 양식도 다름.
  - **작업일보**: 현장 진척 + 수량 데이터. `excel_parser.py`가 처리하는 표준 양식 (C2=공사명, C3=일자, B7=수량텍스트, E38~E42=인원).
  - **업무일지**: 일일 업무 to-do 리스트 (자연어, 수량 없음). 트래커에 자동 반영 불가.
- 과거에 `.env`의 `SUBJECT_KEYWORD=업무일지`로 잘못 설정되어 진짜 작업일보를 못 잡고 있었음.
  2026-05-14 `SUBJECT_KEYWORD=작업일보`로 정정.

### 7.2 다우오피스 대용량 첨부 (.zip) 처리
- 평소 작업일보는 IMAP으로 정상 첨부되지만, 오늘처럼 큰 누적분(70MB 등)은
  다우오피스가 **대용량 첨부 링크**로 전환한다. IMAP에서는 본문에 다운로드 안내문만 보이고
  실제 .zip은 인증된 웹 세션으로만 받을 수 있음 → **자동 처리 불가**.
- 처리 흐름: 사용자가 다우오피스 웹에서 직접 .zip 다운 → `auto-tracker/오늘_첨부/`에 압축 해제 →
  `python batch_import_attachments.py` 실행 → 트래커 반영 → `python make_dashboard.py`.

### 7.3 normalize_line() 정규화 규칙
`excel_parser.normalize_line()` (2026-05-14 강화):
- `상단부`/`상단` → `#1`, `하단부`/`하단` → `#2` (한 라인을 상하단으로 부르는 케이스. ECH2FC 사례.)
- `PBS` → `PB`, `CBS` → `CB`
- 입력에 등장하는 모든 `line` 토큰을 제거한 뒤 끝에 한 번만 붙임 (Line이 중간에 끼는 양식 대응).
- 매핑 예: `PBS Line 상단부` → `pb#1line` → ROW_MAPPING의 `pb#1line_gapcover` 매칭.

### 7.4 파일명 날짜 우선 정책
- 작업자가 파일 복사 후 C3(작업일자) 갱신을 깜빡하는 케이스가 실재함 (2026-05-14 ECH2 사례에서 4/28 파일을 4/29, 4/30용으로 복사).
- 그래서 `parse_attachment()`는 **파일명에서 추출한 날짜를 우선**하고, C3과 불일치하면 WARNING만 띄움.
- 파일명 패턴: `..._26.04.29.xlsm` 등에서 마지막 `YY.MM.DD` 추출.

### 7.5 한글 콘솔 출력 깨짐
PowerShell·Git Bash 콘솔에서 한글이 mojibake로 보여도 파일·엑셀에는 정상 UTF-8로 저장된다. 디버깅 시 콘솔만 보고 판단하지 말 것.
파이썬 스크립트에서 `—`(U+2014) 같은 BMP 외 특수문자는 cp949 콘솔에서 `UnicodeEncodeError` 유발 — `-`로 대체 권장.

### 7.6 파일 잠금
사용자가 엑셀에서 `Mail_Log.xlsx`/`HCNSSERVICE_Tracking_v2.xlsx`를 열고 있으면 `PermissionError`.
- Mail_Log는 자동으로 `*_locked_YYYYMMDD_HHMMSS.xlsx` 사본에 쓴다.
- 트래커는 잠금 시 그 메일/그 파일만 실패 처리 후 다음 계속.
- `batch_import_attachments.py`는 첫 파일에서 잠금 감지 시 즉시 중단(반쪽 머지 방지).
- `merge_tracker_from_email.py`는 비교는 임시 사본으로 가능, 적용은 잠금 풀려야 함.

### 7.7 IMAP SEARCH는 날짜 단위
시간 단위 필터링은 메일 헤더의 `Date`를 fetch 후 코드에서 한 번 더 거른다 (`mail_log_fetcher.search_uids_since`).


---

## 8. 작업 컨벤션

- **언어:** 사용자 대화·로그·메시지·주석 모두 한국어. 영문은 변수명·API 식별자에만.
- **시각:** Asia/Seoul (KST). 로컬 시간으로 저장하고 표기.
- **백업:** 트래커 업데이트는 항상 `backups/` 폴더에 시각 포함 사본 자동 생성 (`excel_updater.make_backup`).
- **민감정보:** EMAIL_PASSWORD, ANTHROPIC_API_KEY는 stdout/로그 어디에도 출력 금지.
- **추정값 표시:** 분류기가 추정한 값은 비고 컬럼에 `추정: <필드>` 또는 자동 생성 메타(`트래커 반영: SITE date N셀`)를 누적.
- **검토필요 우선:** 분류 신뢰도 낮거나 API 실패면 메일을 버리지 말고 `상태=검토필요`로 기록 (스펙 준수).
- **새 모듈 추가 시:** 본 CLAUDE.md의 §3 파일 역할, §2 흐름도, §11 변경 로그에 한 줄 반영.

---

## 9. 자주 쓰는 CLI

```cmd
cd "C:\Users\jyu10\한국씨엔에스(주)\Dashboard_Development - Dashboard\auto-tracker"

REM 파이프라인 ① 1회 (오늘 기준)
python email_fetcher.py
python email_fetcher.py --days 1 --dry-run     REM 어제까지 한 번 미리보기
python email_fetcher.py --days 7               REM 최근 7일 백필

REM 파이프라인 ② 1회 (last_success_ts 이후)
python mail_log_fetcher.py
python mail_log_fetcher.py --since-hours 48    REM 강제 48시간
python mail_log_fetcher.py --dry-run           REM 분류만, 저장 안 함
python mail_log_fetcher.py --no-api            REM Claude 없이 모두 검토필요

REM 대시보드 단독 재생성
python make_dashboard.py            REM 작업 트래커용
python mail_dashboard.py            REM 메일 로그용

REM Mail_Log.xlsx 빈 템플릿 생성 (이미 있으면 skip)
python mail_log_schema.py
python mail_log_schema.py --force   REM 덮어쓰기 (데이터 손실)

REM 수동 첨부 일괄 처리 (대용량 첨부 .zip을 푼 후)
python batch_import_attachments.py --dry-run        REM 확인
python batch_import_attachments.py                  REM 실제 적용
python batch_import_attachments.py --dir 다른_폴더    REM 다른 폴더

REM 이메일로 받은 트래커 사본 안전 머지 (보류 상태)
python merge_tracker_from_email.py <첨부.xlsx>        REM dry-run
python merge_tracker_from_email.py <첨부.xlsx> --apply REM 적용
```

---

## 10. 의존성

`requirements.txt`:
- `openpyxl==3.1.5` — 모든 .xlsx I/O
- `pyperclip==1.9.0` — 카톡 페이스트 (구 경로)
- `anthropic>=0.40.0` — 파이프라인 ② 분류

---

## 11. 변경 로그 (대규모 변경만)

| 날짜 | 변경 | 비고 |
|---|---|---|
| 2026-05-14 | 파이프라인 ② (`mail_log_fetcher.py` 외 4개 모듈) 신설 | Claude Sonnet 4.6 분류, 20필드 추출 |
| 2026-05-14 | 스케줄러 단일 09:00·09:05 → 08:00~10:00·08:05~10:05 30분 간격 | 매일 5회 실행 |
| 2026-05-14 | 파이프라인 ②에 작업일보 첨부 자동 반영 통합 | `try_apply_to_tracker()` |
| 2026-05-14 | 본 CLAUDE.md 작성 — 공유 작업/협업용 컨텍스트 문서로 운영 시작 | |
| 2026-05-14 | `.env` `SUBJECT_KEYWORD` 정정: `업무일지` → `작업일보` | 진짜 작업일보 매칭 |
| 2026-05-14 | `normalize_line()` 정규화 강화 (상단/하단 → #1/#2, line 토큰 위치 보정) | ECH2FC 등 대응 |
| 2026-05-14 | `parse_attachment()` 파일명 날짜 우선 정책 도입 | 작성자 C3 미갱신 흡수 |
| 2026-05-14 | `batch_import_attachments.py` 신설 — 대용량 첨부 등 수동 일괄 처리 | `오늘_첨부/` 폴더 |
| 2026-05-14 | `merge_tracker_from_email.py` 신설 — 트래커 사본 안전 머지 (보류) | 레이아웃 차이로 미적용 |
| 2026-05-14 | 백필 1회 적용: 4/9, 4/21~30 작업일보 14개 → 트래커 69셀 업데이트 | CHW1FC + ECH2FC |

---

## 12. 다음에 누가 이어받을 때 가장 먼저 볼 것

1. 본 CLAUDE.md 전체 (5분).
2. `.env`가 채워져 있는지 (`ANTHROPIC_API_KEY` 비어있을 수 있음 — 그러면 검토필요로 쌓임).
3. 스케줄러에 두 작업이 등록돼 있는지 (`schtasks /query | findstr HCNS`).
4. `Mail_Log.xlsx`와 `HCNSSERVICE_Tracking_v2.xlsx` 잠금 상태 (엑셀로 열어두면 안 됨).
5. §7 "알려진 이슈" — 특히 7.1 양식 호환성. 현재 IMAP에 도착하는 첨부는 트래커 자동 반영 안 됨.
