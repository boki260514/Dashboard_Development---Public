"""
다우오피스 메일 → Mail_Log.xlsx + mail_dashboard.html 자동 업데이트.

매일 오전 9:00(KST) Windows 작업 스케줄러에서 실행되며,
마지막 성공 실행 시각 이후 도착한 메일을 처리한다.

흐름:
    1) .env 로드 (IMAP/SMTP/ANTHROPIC_API_KEY)
    2) mail_state.json에서 last_success_ts 로딩 — 없으면 24시간 전
    3) IMAP INBOX에서 last_success_ts 이후 메일 UID 검색
    4) 각 메일을 광고 1차 필터 → Claude API 분류 → Mail_Log 후보로 변환
    5) 기존 Mail_Log.xlsx의 메일ID/fingerprint와 비교해 중복 제외
    6) Mail_Log.xlsx에 신규 행 append + Run_Log에 1행 기록
    7) mail_dashboard.html 재생성
    8) mail_state.json 업데이트 (이번 실행 종료 시각)
    9) 표준 출력으로 결과 보고

오류 처리:
    - IMAP/엑셀/대시보드 어디서 실패해도 가능한 범위까지 처리.
    - 모든 오류는 Run_Log의 오류상세에 기록됨.

사용:
    python mail_log_fetcher.py                # 평소 운영
    python mail_log_fetcher.py --since-hours 48  # 최근 48시간만 (백필)
    python mail_log_fetcher.py --dry-run       # 분류는 하되 엑셀/대시보드 저장 안 함
    python mail_log_fetcher.py --no-api        # Claude API 호출 안 함 (모든 메일 검토필요)
"""
from __future__ import annotations

import argparse
import email
import imaplib
import json
import logging
import os
import re
import sys
import tempfile
import traceback
from datetime import date, datetime, timedelta, timezone
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mail_classifier import classify_mail
from mail_log_excel import (
    MailRow, RunLogRow,
    append_rows, ensure_workbook, load_existing_keys, make_fingerprint,
)
from mail_log_schema import MAIL_LOG_XLSX, MAIL_STATE_JSON
import mail_dashboard

# 기존 작업일보 트래커 연동 모듈
from config import EXCEL_PATH as TRACKER_XLSX
from excel_parser import parse_attachment as parse_tracker_attachment
from excel_updater import update_excel as update_tracker_excel


# ============================================================
# 로깅
# ============================================================
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt="%H:%M:%S")
logger = logging.getLogger("mail_log_fetcher")


# ============================================================
# .env 로더 (email_fetcher.py와 동일 형식)
# ============================================================
def load_env(env_path: Optional[Path] = None) -> dict:
    env_path = env_path or (Path(__file__).parent / ".env")
    data: dict[str, str] = {}
    if not env_path.exists():
        return data
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k, v = k.strip(), v.strip()
        if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
            v = v[1:-1]
        data[k] = v
    return data


# ============================================================
# 상태 파일
# ============================================================
def load_state() -> dict:
    if not MAIL_STATE_JSON.exists():
        return {}
    try:
        return json.loads(MAIL_STATE_JSON.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("mail_state.json 파싱 실패: %s", exc)
        return {}


def save_state(state: dict) -> None:
    MAIL_STATE_JSON.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def next_run_id(state: dict) -> str:
    seq = int(state.get("run_seq", 0)) + 1
    return f"RUN-{datetime.now().strftime('%Y%m%d')}-{seq:03d}"


# ============================================================
# 메일 헤더/본문 디코딩
# ============================================================
def decode_text(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def extract_body_text(msg: email.message.Message) -> str:
    """text/plain 우선, 없으면 text/html에서 태그 제거."""
    plain = []
    html_text = []
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = (part.get("Content-Disposition") or "").lower()
        if "attachment" in disp:
            continue
        if ctype == "text/plain":
            try:
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                plain.append(payload.decode(charset, errors="replace"))
            except Exception:
                continue
        elif ctype == "text/html":
            try:
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                html_text.append(payload.decode(charset, errors="replace"))
            except Exception:
                continue
    if plain:
        return "\n\n".join(plain).strip()
    if html_text:
        # 매우 단순한 태그 제거 (의존성 추가 없이)
        joined = "\n\n".join(html_text)
        joined = re.sub(r"<\s*(script|style)[^>]*>.*?<\s*/\s*\1\s*>", "", joined,
                        flags=re.IGNORECASE | re.DOTALL)
        joined = re.sub(r"<[^>]+>", " ", joined)
        joined = re.sub(r"\s+", " ", joined)
        return joined.strip()
    return ""


def list_attachment_names(msg: email.message.Message) -> list[str]:
    names: list[str] = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if filename:
            names.append(decode_text(filename))
    return names


def save_attachments(msg: email.message.Message, save_dir: Path) -> list[Path]:
    """메시지의 모든 첨부를 save_dir에 저장하고 경로 리스트 반환.

    파일명 충돌 방지: 같은 이름이 있으면 _1, _2 ... 접미사를 붙임.
    """
    saved: list[Path] = []
    used_names: set[str] = set()
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if not filename:
            continue
        filename = decode_text(filename)
        # 파일명 충돌 회피
        target_name = filename
        stem, dot, ext = filename.rpartition(".")
        i = 1
        while target_name in used_names:
            if dot:
                target_name = f"{stem}_{i}.{ext}"
            else:
                target_name = f"{filename}_{i}"
            i += 1
        used_names.add(target_name)
        try:
            payload = part.get_payload(decode=True) or b""
            target = save_dir / target_name
            target.write_bytes(payload)
            saved.append(target)
        except Exception as exc:
            logger.warning("첨부 저장 실패 (%s): %s", filename, exc)
    return saved


def try_apply_to_tracker(
    attachment_paths: list[Path],
) -> tuple[str, Optional[str], list[str]]:
    """첨부 중 작업일보 형식 엑셀이 있으면 HCNSSERVICE 트래커에 반영한다.

    Returns:
        (요약문, 백업경로 or None, 에러 메시지 리스트)
        - 작업일보 형식 첨부가 없으면 ("", None, [])
        - 파싱은 됐지만 site/date 추출 실패면 ("트래커 미반영: 형식 비매칭", None, [])
        - 성공 시 ("트래커 반영: SITE date N셀", backup_path, [])
    """
    excel_atts = [p for p in attachment_paths
                  if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    if not excel_atts:
        return "", None, []

    summaries: list[str] = []
    errors: list[str] = []
    backup_path: Optional[str] = None

    for att in excel_atts:
        try:
            report = parse_tracker_attachment(att)
        except Exception as exc:
            errors.append(f"{att.name}: 파싱 예외 {exc}")
            continue
        if not report or not report.site_sheet or not report.work_date:
            # 작업일보 형식이 아닌 일반 엑셀(견적서/주문서 등) — 정상 케이스
            summaries.append(f"{att.name}: 트래커 미반영(형식 비매칭)")
            continue
        try:
            result = update_tracker_excel(
                report, TRACKER_XLSX, create_backup=True, overwrite=True,
            )
            cells = len(result.results)
            errs = len(result.errors)
            skips = len(result.skipped)
            summaries.append(
                f"{att.name}: 트래커 {report.site_sheet} {report.work_date} "
                f"({cells}셀{', '+str(errs)+'에러' if errs else ''}"
                f"{', '+str(skips)+'스킵' if skips else ''})"
            )
            if result.backup_path and not backup_path:
                backup_path = result.backup_path
            for e in result.errors:
                errors.append(f"{att.name}: {e}")
        except PermissionError as exc:
            errors.append(f"{att.name}: 트래커 파일 잠김 - {exc}")
        except Exception as exc:
            errors.append(f"{att.name}: 트래커 업데이트 실패 - {exc}")

    return " / ".join(summaries), backup_path, errors


# ============================================================
# IMAP 메일 검색
# ============================================================
def search_uids_since(imap: imaplib.IMAP4_SSL, since_dt: datetime) -> list[bytes]:
    """since_dt 이후 도착한 메일의 UID 목록.

    IMAP SEARCH의 SINCE는 날짜 단위라서 시간 비교는 fetch 후 직접 한다.
    """
    imap.select("INBOX", readonly=True)
    since_date_str = since_dt.strftime("%d-%b-%Y")
    try:
        typ, data = imap.uid("SEARCH", None, "SINCE", since_date_str)
    except imaplib.IMAP4.error as exc:
        logger.error("IMAP SEARCH 실패: %s", exc)
        return []
    if typ != "OK" or not data or not data[0]:
        return []
    return data[0].split()


def fetch_message(imap: imaplib.IMAP4_SSL, uid: bytes) -> Optional[email.message.Message]:
    for arg in ("(BODY.PEEK[])", "(RFC822.PEEK)", "(RFC822)"):
        try:
            typ, data = imap.uid("FETCH", uid, arg)
        except imaplib.IMAP4.error:
            continue
        if typ != "OK" or not data:
            continue
        for part in data:
            if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)) and len(part[1]) > 50:
                try:
                    return email.message_from_bytes(bytes(part[1]))
                except Exception:
                    continue
    return None


# ============================================================
# 메인
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--since-hours", type=float, default=None,
                    help="last_success_ts 무시하고 최근 N시간으로 강제")
    ap.add_argument("--dry-run", action="store_true",
                    help="엑셀/대시보드 저장 안 함")
    ap.add_argument("--no-api", action="store_true",
                    help="Claude API 호출 없이 모두 검토필요로 기록")
    args = ap.parse_args()

    started_at = datetime.now()
    run_started_str = started_at.strftime("%Y-%m-%d %H:%M:%S")

    # 1) .env
    env = load_env()
    required_imap = ["IMAP_HOST", "IMAP_PORT", "EMAIL_ADDRESS", "EMAIL_PASSWORD"]
    missing = [k for k in required_imap if not env.get(k)]
    if missing:
        logger.error(".env 누락: %s", ", ".join(missing))
        return 1
    imap_host = env["IMAP_HOST"]
    imap_port = int(env["IMAP_PORT"])
    email_addr = env["EMAIL_ADDRESS"]
    email_pass = env["EMAIL_PASSWORD"]
    api_key = env.get("ANTHROPIC_API_KEY", "") or os.environ.get("ANTHROPIC_API_KEY", "")
    if args.no_api:
        api_key = ""

    # 2) 상태 로드 & since 결정
    state = load_state()
    if args.since_hours is not None:
        since_dt = datetime.now() - timedelta(hours=args.since_hours)
    else:
        last = state.get("last_success_ts")
        if last:
            try:
                since_dt = datetime.fromisoformat(last)
            except ValueError:
                since_dt = datetime.now() - timedelta(hours=24)
        else:
            since_dt = datetime.now() - timedelta(hours=24)
    logger.info("처리 범위: %s 이후", since_dt.strftime("%Y-%m-%d %H:%M:%S"))

    # 3) Mail_Log.xlsx 보장 & 기존 키 로드
    ensure_workbook(MAIL_LOG_XLSX)
    existing_ids, existing_fps = load_existing_keys(MAIL_LOG_XLSX)
    logger.info("기존 메일ID %d개 / fingerprint %d개", len(existing_ids), len(existing_fps))

    # 4) IMAP
    try:
        imap = imaplib.IMAP4_SSL(imap_host, imap_port, timeout=30)
        imap.login(email_addr, email_pass)
    except Exception as exc:
        logger.error("IMAP 로그인 실패: %s", exc)
        _write_failed_runlog(state, run_started_str, "IMAP 로그인 실패", args.dry_run)
        return 2

    uids = search_uids_since(imap, since_dt)
    logger.info("UID 검색 결과: %d개", len(uids))

    checked = 0
    new_rows: list[MailRow] = []
    dup_count = 0
    review_count = 0
    non_business_count = 0
    tracker_updated_count = 0
    error_details: list[str] = []

    # 첨부 다운로드용 임시 디렉터리 (메일 루프 전체에서 1개 사용 후 정리)
    with tempfile.TemporaryDirectory(prefix="mail_log_") as _tmp:
        tmp_dir = Path(_tmp)

        for uid in uids:
            msg = fetch_message(imap, uid)
            if msg is None:
                continue
            checked += 1

            subject = decode_text(msg.get("Subject", ""))
            from_name, from_addr = parseaddr(msg.get("From", ""))
            from_name = decode_text(from_name) or (from_name or "").strip()
            from_addr = (from_addr or "").lower().strip()

            recv_dt: Optional[datetime] = None
            date_hdr = msg.get("Date")
            if date_hdr:
                try:
                    recv_dt = parsedate_to_datetime(date_hdr)
                    if recv_dt.tzinfo:
                        recv_dt = recv_dt.astimezone().replace(tzinfo=None)
                except Exception:
                    recv_dt = None
            if recv_dt is None:
                recv_dt = datetime.now()
            if recv_dt < since_dt:
                continue

            recv_str = recv_dt.strftime("%Y-%m-%d %H:%M:%S")

            msg_id = (msg.get("Message-ID") or "").strip()
            mail_id = msg_id or f"UID-{uid.decode(errors='replace')}"

            # 첨부: 이름만 먼저 확인 (중복이면 다운로드 안 함)
            att_names = list_attachment_names(msg)
            att_str = ", ".join(att_names)

            fp = make_fingerprint(recv_str, from_addr, subject)
            if mail_id in existing_ids or fp in existing_fps:
                dup_count += 1
                continue

            # 첨부 실제 저장 (메일ID 단위 폴더)
            uid_dir = tmp_dir / re.sub(r"[^A-Za-z0-9_-]", "_", mail_id)[:40]
            uid_dir.mkdir(parents=True, exist_ok=True)
            att_paths = save_attachments(msg, uid_dir)

            body = extract_body_text(msg)

            try:
                result = classify_mail(
                    api_key=api_key,
                    sender_name=from_name,
                    sender_email=from_addr,
                    subject=subject,
                    received_at=recv_str,
                    body_text=body,
                    attachments=att_names,
                )
            except Exception as exc:
                logger.error("분류 중 예외 (UID=%s): %s",
                             uid.decode(errors="replace"), exc)
                error_details.append(
                    f"UID={uid.decode(errors='replace')}: 분류 예외 {exc}")
                result = None

            # 첨부 엑셀이 작업일보 형식이면 HCNSSERVICE 트래커에 반영
            tracker_summary = ""
            if att_paths:
                try:
                    tracker_summary, _bk, tracker_errs = try_apply_to_tracker(att_paths)
                    if tracker_errs:
                        error_details.extend(tracker_errs)
                    if tracker_summary and "트래커" in tracker_summary and "셀" in tracker_summary:
                        tracker_updated_count += 1
                except Exception as exc:
                    logger.error("트래커 반영 예외: %s", exc)
                    error_details.append(
                        f"UID={uid.decode(errors='replace')}: 트래커 반영 예외 {exc}")

            def _merge_note(*parts: str) -> str:
                return " | ".join(p for p in parts if p)

            if result is None:
                row = MailRow(
                    mail_id=mail_id, received_at=recv_str,
                    sender_name=from_name, sender_email=from_addr,
                    company="", subject=subject,
                    work_type="기타", summary="", request="",
                    amount="", quantity="", due_date="",
                    owner="", priority="보통", status="검토필요",
                    attachments=att_str, ref=uid.decode(errors="replace"),
                    processed_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    confidence="낮음",
                    note=_merge_note("자동분류 실패", tracker_summary),
                )
                new_rows.append(row)
                review_count += 1
                existing_ids.add(mail_id)
                existing_fps.add(fp)
                continue

            if not result.is_business:
                # 비업무로 분류됐어도 트래커에 반영된 첨부는 정상 케이스
                non_business_count += 1
                existing_ids.add(mail_id)
                existing_fps.add(fp)
                continue

            f = result.fields
            row = MailRow(
                mail_id=mail_id, received_at=recv_str,
                sender_name=from_name, sender_email=from_addr,
                company=f.get("회사명", ""), subject=subject,
                work_type=f.get("업무유형", "기타"),
                summary=f.get("핵심내용요약", ""),
                request=f.get("요청사항", ""),
                amount=f.get("금액", ""),
                quantity=f.get("수량", ""),
                due_date=f.get("납기일", ""),
                owner=f.get("담당자", ""),
                priority=f.get("우선순위", "보통"),
                status=f.get("상태", "신규"),
                attachments=att_str, ref=uid.decode(errors="replace"),
                processed_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                confidence=f.get("신뢰도", "중간"),
                note=_merge_note(f.get("비고", ""), tracker_summary),
            )
            new_rows.append(row)
            if row.status == "검토필요":
                review_count += 1
            existing_ids.add(mail_id)
            existing_fps.add(fp)

    try:
        imap.logout()
    except Exception:
        pass

    # 5) 엑셀 append + Run_Log
    ended_at = datetime.now()
    run_id = next_run_id(state)
    dashboard_status = "skipped"
    appended = 0
    fallback_path = None

    has_error = bool(error_details)
    if args.dry_run:
        logger.info("[DRY-RUN] 엑셀 저장 생략 — 신규 후보 %d건", len(new_rows))
    else:
        try:
            run_log = RunLogRow(
                run_id=run_id,
                started_at=run_started_str,
                ended_at=ended_at.strftime("%Y-%m-%d %H:%M:%S"),
                checked=checked,
                new_count=len(new_rows),
                duplicate=dup_count,
                review_needed=review_count,
                non_business=non_business_count,
                has_error=has_error,
                error_detail=" | ".join(error_details)[:500],
                dashboard_status="pending",
            )
            appended, fallback_path = append_rows(new_rows, run_log=run_log)
        except Exception as exc:
            logger.error("엑셀 저장 실패: %s", exc)
            error_details.append(f"엑셀 저장 실패: {exc}")
            has_error = True

    # 6) 대시보드 재생성
    if not args.dry_run:
        try:
            mail_dashboard.regenerate(quiet=True)
            dashboard_status = "ok"
        except Exception as exc:
            logger.error("대시보드 재생성 실패: %s", exc)
            error_details.append(f"대시보드 재생성 실패: {exc}")
            dashboard_status = "failed"

    # 7) 상태 업데이트 — 오류와 무관하게 진행한 만큼 다음 컷오프를 갱신
    state["last_success_ts"] = ended_at.isoformat(timespec="seconds")
    state["last_run_id"] = run_id
    state["run_seq"] = int(state.get("run_seq", 0)) + 1
    if not args.dry_run:
        save_state(state)

    # 8) 최종 보고
    _print_report(
        started_at=started_at, ended_at=ended_at,
        checked=checked, new_count=appended if not args.dry_run else len(new_rows),
        dup_count=dup_count, review_count=review_count,
        non_business=non_business_count,
        tracker_updated=tracker_updated_count,
        dashboard_status=dashboard_status,
        has_error=has_error, errors=error_details,
        new_rows=new_rows, fallback_path=fallback_path,
        dry_run=args.dry_run,
    )
    return 0 if not has_error else 3


def _write_failed_runlog(state: dict, started_str: str, reason: str, dry_run: bool) -> None:
    if dry_run:
        return
    try:
        run_log = RunLogRow(
            run_id=next_run_id(state),
            started_at=started_str,
            ended_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            checked=0, new_count=0, duplicate=0, review_needed=0,
            non_business=0,
            has_error=True, error_detail=reason,
            dashboard_status="skipped",
        )
        append_rows([], run_log=run_log)
    except Exception:
        pass


def _print_report(
    *, started_at, ended_at, checked, new_count, dup_count, review_count,
    non_business, tracker_updated, dashboard_status, has_error, errors,
    new_rows, fallback_path, dry_run,
):
    print()
    print("=" * 60)
    print("[일일 메일 자동 업데이트 결과]")
    print("=" * 60)
    print(f"실행 일시: {started_at.strftime('%Y-%m-%d %H:%M:%S')} ~ {ended_at.strftime('%H:%M:%S')}")
    print(f"확인한 메일 수: {checked}")
    print(f"신규 반영 건수: {new_count}{' (DRY-RUN)' if dry_run else ''}")
    print(f"중복 제외 건수: {dup_count}")
    print(f"검토필요 건수: {review_count}")
    print(f"비업무 제외 건수: {non_business}")
    print(f"트래커(HCNSSERVICE_Tracking_v2) 반영 건수: {tracker_updated}")
    print(f"대시보드 새로고침 상태: {dashboard_status}")
    print(f"오류 여부: {'Y' if has_error else 'N'}")
    if fallback_path:
        print(f"⚠️ 원본 잠금으로 임시 사본에 기록: {fallback_path}")
    if errors:
        print()
        print("오류 상세:")
        for e in errors[:5]:
            print(f"  - {e}")

    # 주요 신규 항목 — 우선순위 높은 순
    biz_rows = [r for r in new_rows if r.status != "검토필요"]
    pri_rank = {"긴급": 0, "높음": 1, "보통": 2, "낮음": 3}
    biz_rows.sort(key=lambda r: pri_rank.get(r.priority, 9))

    print()
    print("주요 신규 항목 요약:")
    if not biz_rows:
        print("  (없음)")
    for i, r in enumerate(biz_rows[:5], start=1):
        amt = f" / 금액 {int(r.amount):,}원" if isinstance(r.amount, (int, float)) and r.amount else ""
        due = f" / 납기 {r.due_date}" if r.due_date else ""
        print(f"  {i}. [{r.priority}/{r.work_type}] {r.company or '-'} | {r.subject[:50]}{amt}{due}")

    print()
    print("검토가 필요한 항목:")
    review_rows = [r for r in new_rows if r.status == "검토필요"]
    if not review_rows:
        print("  (없음)")
    for i, r in enumerate(review_rows[:5], start=1):
        print(f"  {i}. [{r.priority}] {r.sender_email} | {r.subject[:50]} ({r.note[:30]})")

    print()
    print("다음 액션:")
    if review_rows:
        print(f"  - Mail_Log.xlsx에서 상태='검토필요' 행 {len(review_rows)}건 확인 필요")
    urgent = [r for r in biz_rows if r.priority in ("긴급", "높음")]
    if urgent:
        print(f"  - 긴급/높음 우선순위 {len(urgent)}건 즉시 처리 검토")
    if not review_rows and not urgent:
        print("  - 추가 액션 없음")
    print("=" * 60)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        logger.error("사용자 중단")
        sys.exit(130)
    except Exception:
        logger.error("치명적 오류 발생:\n%s", traceback.format_exc())
        sys.exit(99)
