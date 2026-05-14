# 📧 일반 업무 메일 로그 자동화

다우오피스 INBOX의 일반 업무 메일을 매일 오전 8:05~10:05 (30분 간격, 5회)에 자동으로
**Mail_Log.xlsx** 에 누적하고 **mail_dashboard.html** 을 갱신합니다.

기존 작업일보 자동화(`email_fetcher.py`, 09:00 실행)와는 별개로 동작합니다.

---

## 처음 한 번 (설치)

### 1. 의존성 설치

```cmd
cd "C:\Users\jyu10\한국씨엔에스(주)\Dashboard_Development - Dashboard\auto-tracker"
pip install -r requirements.txt
```

`anthropic` 패키지가 추가로 설치됩니다.

### 2. Claude API 키 등록

`.env` 파일을 열어서 비어 있는 `ANTHROPIC_API_KEY=` 줄에 본인 API 키를 넣습니다.

```
ANTHROPIC_API_KEY=sk-ant-api03-...
```

키는 https://console.anthropic.com/ 의 **API Keys** 메뉴에서 발급합니다.

> 키를 비워두고 실행해도 시스템은 동작합니다. 다만 모든 메일이
> 상태=`검토필요`, 신뢰도=`낮음`으로 기록되어 사람이 직접 분류해야 합니다.

### 3. Mail_Log.xlsx 미리 만들기 (선택)

자동으로도 생성되지만 미리 만들어 두려면:

```cmd
python mail_log_schema.py
```

→ 상위 폴더(`Dashboard_Development - Dashboard`)에 `Mail_Log.xlsx`가 생성됩니다.

### 4. 작업 스케줄러에 등록

`자동화_등록_메일로그.bat` 을 **오른쪽 클릭 → 관리자 권한으로 실행**.

매일 오전 **9:05** 에 자동으로 실행됩니다.
(작업일보 자동화가 09:00이라 5분 어긋나게 등록)

---

## 매일 자동으로 일어나는 일

매일 09:05에 다음이 순차로 실행됩니다.

1. `.env`의 IMAP 정보로 다우오피스 INBOX 접속
2. `mail_state.json`의 마지막 성공 시각 이후 도착한 메일 검색
   - 상태 파일이 없으면 **최근 24시간**
3. 광고/뉴스레터/시스템 자동발신 메일은 1차 필터로 제외 (API 호출 전)
4. 남은 메일을 한 통씩 Claude Sonnet 4.6에 보내 20개 필드 추출
5. **첨부에 작업일보 형식 엑셀이 있으면** `excel_parser` 로 파싱해
   `HCNSSERVICE_Tracking_v2.xlsx` 에 자동 반영
   (실패해도 메일 로그 자체는 계속 진행. 비고에 반영 결과 기록.)
6. 기존 `Mail_Log.xlsx` 와 비교해 중복 제외 (메일ID 우선)
7. 신규 행을 `Mail_Log.xlsx` 의 **Mail_Log** 시트에 append
8. `Run_Log` 시트에 이번 실행 통계 1줄 추가
9. `mail_dashboard.html` 재생성
10. 상태 파일(`mail_state.json`)을 종료 시각으로 업데이트

---

## 파일 구조

```
auto-tracker/
├── mail_log_fetcher.py       ← 메인 (스케줄러가 실행)
├── mail_classifier.py        ← Claude API 호출 + JSON 파싱
├── mail_log_excel.py         ← Mail_Log.xlsx 읽기/쓰기 (중복 체크)
├── mail_log_schema.py        ← 컬럼 정의 + 빈 템플릿 생성기
├── mail_dashboard.py         ← mail_dashboard.html 생성기
├── mail_dashboard.html       ← 결과 대시보드 ✨
├── mail_state.json           ← 마지막 성공 실행 시각 (자동 관리)
├── 메일로그처리.bat            ← 수동 실행용
├── 자동화_등록_메일로그.bat    ← 작업 스케줄러 등록
└── 자동화_해제_메일로그.bat    ← 해제

../
└── Mail_Log.xlsx             ← 누적 데이터 (Dashboard_Development - Dashboard 폴더)
```

---

## Mail_Log.xlsx 컬럼 (20개)

| # | 컬럼 | 형식/값 |
|---:|---|---|
| 1  | 메일ID         | `Message-ID` 또는 `UID-xxx` |
| 2  | 수신일시       | `YYYY-MM-DD HH:mm:ss` |
| 3  | 발신자이름     | 텍스트 |
| 4  | 발신자이메일   | 텍스트 |
| 5  | 회사명         | 발신자 추정 |
| 6  | 제목           | 메일 제목 원문 |
| 7  | 업무유형       | 견적/주문/계약/정산/CS/일정/보고/기타 |
| 8  | 핵심내용요약   | 1~3문장 |
| 9  | 요청사항       | 발신자가 요구하는 액션 |
| 10 | 금액           | 숫자(원). 없으면 빈칸 |
| 11 | 수량           | 숫자. 없으면 빈칸 |
| 12 | 납기일         | `YYYY-MM-DD`. 없으면 빈칸 |
| 13 | 담당자         | 본문/서명에서 추정 |
| 14 | 우선순위       | 긴급/높음/보통/낮음 |
| 15 | 상태           | 신규/진행중/검토필요/완료 |
| 16 | 첨부파일명     | 콤마 구분 |
| 17 | 원문참조       | IMAP UID |
| 18 | 처리일시       | `YYYY-MM-DD HH:mm:ss` |
| 19 | 신뢰도         | 높음/중간/낮음 |
| 20 | 비고           | 추정값 표기 등 |

### 사람이 직접 수정해도 되나요?

- **상태**, **담당자**, **회사명**, **비고** 같은 컬럼은 사람이 직접 수정 가능합니다.
- 다음 자동 실행은 **메일ID 기준**으로 중복을 제외하므로, **기존 행은 덮어쓰지 않습니다**.
- 즉, 한 번 등록된 메일을 사람이 수정해도 자동화가 다시 건드리지 않습니다.

---

## 수동 실행 (운영 중)

```cmd
cd "C:\Users\jyu10\한국씨엔에스(주)\Dashboard_Development - Dashboard\auto-tracker"

REM 평소 운영 (last_success_ts 이후)
python mail_log_fetcher.py

REM 최근 48시간 강제 백필
python mail_log_fetcher.py --since-hours 48

REM 분류는 하되 엑셀/대시보드에 쓰지 않음 (테스트용)
python mail_log_fetcher.py --since-hours 24 --dry-run

REM Claude API 없이 모두 검토필요로 기록
python mail_log_fetcher.py --no-api
```

---

## 대시보드 표시 지표 (7개)

`mail_dashboard.html` 을 더블클릭하면 브라우저에서 열립니다.

1. 최근 7일 신규 메일 건수
2. 긴급+높음 우선순위 건수
3. 검토필요 건수
4. 마감일 임박(≤7일) 건수
5. 금액 합계 (전체 + 최근 7일)
6. 업무 유형별 건수
7. 고객사 TOP 10 + 최근 메일 20건 테이블

---

## 오류·예외 처리

스펙대로 **작업을 중단하지 않고 가능한 범위까지** 처리합니다.

| 상황 | 처리 |
|---|---|
| IMAP 로그인 실패 | Run_Log에 실패로 기록 후 종료 |
| Claude API 호출 실패 | 그 메일만 상태=`검토필요`로 저장 후 다음 메일 계속 |
| 본문 디코딩 실패 | 빈 본문으로 분류 (대부분 검토필요로 분류됨) |
| Mail_Log.xlsx 잠김(엑셀로 열려있음) | `Mail_Log_locked_yyyymmdd_HHMMSS.xlsx` 임시 사본에 기록 |
| 대시보드 생성 실패 | 엑셀 업데이트는 유지, Run_Log에 `대시보드상태=failed` |
| JSON 파싱 실패 | 그 메일만 검토필요, 다음 메일 계속 |

비밀번호·API 키 등 민감정보는 어떤 로그에도 출력되지 않습니다.

---

## 분류 정확도 끌어올리기

처음 한 주는 **검토필요가 많이 나오는 게 정상**입니다. 다음을 하시면 좋습니다.

1. 처음 며칠은 매일 `Mail_Log.xlsx` 의 `검토필요` 행을 사람이 확인해서
   업무유형/우선순위/상태를 직접 수정합니다.
2. 광고/시스템 발신자 도메인이 자주 통과하면 `mail_classifier.py` 상단의
   `_AD_FROM_PATTERNS` 에 추가합니다.
3. 본인 회사 도메인의 사내 알림이 업무로 잘못 분류되면 같은 위치의
   `_AD_SUBJECT_PATTERNS` 에 사내 알림 제목 키워드를 추가합니다.

---

## 자주 발생하는 문제

| 증상 | 원인 / 해결 |
|---|---|
| `anthropic 패키지 미설치` | `pip install -r requirements.txt` |
| `ANTHROPIC_API_KEY 미설정` | `.env` 의 `ANTHROPIC_API_KEY=` 채우기 |
| `IMAP 로그인 실패` | `.env` 의 EMAIL_ADDRESS/EMAIL_PASSWORD 확인. 비밀번호 변경 직후라면 갱신 |
| Mail_Log.xlsx가 잠겨서 안 써짐 | 엑셀에서 닫고 다시 실행. 못 닫는 상황이면 `_locked_*.xlsx` 사본 확인 |
| 분류가 자꾸 `검토필요`로 나옴 | API 키가 비어있는지, 키 잔액이 있는지 확인 |
| 같은 메일이 또 들어옴 | `mail_state.json` 의 `last_success_ts` 가 갱신되는지 확인 |

---

## 작업일보 자동화와의 관계

| 항목 | 작업일보 자동화 (기존) | 메일 로그 자동화 (NEW) |
|---|---|---|
| 메인 스크립트 | `email_fetcher.py` | `mail_log_fetcher.py` |
| 처리 대상 | 특정 발신자의 '작업일보' 제목 메일 | INBOX 전체의 업무 메일 |
| 첨부 처리 | 엑셀 파싱 후 트래킹 시트 행 단위 업데이트 | 파일명만 기록 (내용 분석 X) |
| 출력 엑셀 | `HCNSSERVICE_Tracking_v2.xlsx` | `Mail_Log.xlsx` (별도 파일) |
| 출력 대시보드 | `dashboard.html` | `mail_dashboard.html` |
| 스케줄 | 매일 08:00~10:00 30분 간격(5회) | 매일 08:05~10:05 30분 간격(5회) |

서로 영향 없이 동시에 운영 가능합니다.
