"""
다우오피스 메일에서 작업일보 첨부파일을 가져와 트래커/대시보드를 자동 업데이트.

매일 오전 정해진 시각에 스케줄링되어 실행됩니다.

동작 흐름:
    1) IMAP으로 메일함 접속
    2) 허용된 발신자 + 제목에 '작업일보' 키워드 포함된 메일 검색 (오늘 자)
    3) 엑셀 첨부파일 다운로드
    4) excel_parser로 파싱 → excel_updater로 트래커 업데이트
    5) make_dashboard로 대시보드 HTML 재생성
    6) 오늘 자 작업일보 메일이 한 통도 없으면 본인에게 알림 메일 발송

설정은 같은 폴더의 .env 파일에서 읽어옵니다.

사용법:
    python email_fetcher.py            # 오늘 메일만 처리 (스케줄용)
    python email_fetcher.py --days N   # 최근 N일치 메일 처리 (수동 백필)
    python email_fetcher.py --dry-run  # 실제 업데이트 안 하고 결과만 출력
"""
from __future__ import annotations

import argparse
import email
import imaplib
import logging
import os
import re
import smtplib
import ssl
import sys
import tempfile
from datetime import date, datetime, timedelta
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Iterable, Optional

# 동일 폴더의 다른 모듈
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_parser import parse_attachment
from excel_updater import update_excel
from config import EXCEL_PATH


# ============================================================
# 로깅
# ============================================================
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt="%H:%M:%S")
logger = logging.getLogger("email_fetcher")


# ============================================================
# .env 파일 로더 (의존성 없음)
# ============================================================
def load_env(env_path: str | os.PathLike = None) -> dict:
    """간단한 KEY=VALUE 형식의 .env 파일 파서."""
    if env_path is None:
        env_path = Path(__file__).parent / ".env"
    env_path = Path(env_path)
    data: dict[str, str] = {}
    if not env_path.exists():
        logger.error(".env 파일이 없습니다: %s", env_path)
        return data
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip()
        # 따옴표 제거
        if (val.startswith('"') and val.endswith('"')) or \
           (val.startswith("'") and val.endswith("'")):
            val = val[1:-1]
        data[key] = val
    return data


# ============================================================
# 메일 헤더 / 디코딩 유틸
# ============================================================
def decode_text(text: Optional[str]) -> str:
    """RFC2047 인코딩된 헤더(이메일 제목 등)를 평문으로 변환."""
    if not text:
        return ""
    try:
        return str(make_header(decode_header(text)))
    except Exception:
        return text


def extract_email_address(from_header: str) -> str:
    """'홍길동 <a@b.com>' → 'a@b.com'"""
    _, addr = parseaddr(from_header or "")
    return (addr or "").lower().strip()


# ============================================================
# IMAP: 작업일보 메일 검색 및 첨부 다운로드
# ============================================================
def search_daily_reports(
    imap: imaplib.IMAP4_SSL,
    allowed_senders: Iterable[str],
    subject_keywords: Iterable[str],
    target_date: date,
) -> list[bytes]:
    """
    오늘 받은(또는 target_date에 받은) 메일 중 허용 발신자 + 제목 키워드 일치하는 메시지 UID 목록.

    Returns:
        IMAP UID(bytes) 리스트.
    """
    imap.select("INBOX", readonly=True)

    # 검색 범위: target_date ~ target_date+1
    since = target_date.strftime("%d-%b-%Y")
    before = (target_date + timedelta(days=1)).strftime("%d-%b-%Y")

    matched: list[bytes] = []
    allowed_set = {a.lower() for a in allowed_senders if a}

    # 발신자별로 SEARCH 실행 (안전한 IMAP 검색)
    for sender in allowed_set:
        # SEARCH FROM "sender" SINCE date BEFORE date
        try:
            typ, data = imap.uid(
                "SEARCH", None,
                "FROM", f'"{sender}"',
                "SINCE", since,
                "BEFORE", before,
            )
        except imaplib.IMAP4.error as exc:
            logger.warning("IMAP SEARCH 실패 (%s): %s", sender, exc)
            continue

        if typ != "OK" or not data or not data[0]:
            continue

        for uid in data[0].split():
            matched.append(uid)

    # 제목 필터링 (다우오피스의 IMAP SEARCH SUBJECT가 한글에 약할 수 있어서
    # 메시지 헤더를 직접 가져와 필터)
    final: list[bytes] = []
    seen = set()
    for uid in matched:
        if uid in seen:
            continue
        seen.add(uid)
        try:
            typ, data = imap.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
        except imaplib.IMAP4.error as exc:
            logger.warning("HEADER FETCH 실패 (UID %s): %s", uid, exc)
            continue
        if typ != "OK" or not data or not data[0]:
            continue
        # data는 리스트, 각 원소는 tuple(header_id, payload)
        payload = data[0][1] if isinstance(data[0], tuple) else b""
        try:
            msg = email.message_from_bytes(payload)
        except Exception:
            continue
        subj = decode_text(msg.get("Subject", ""))
        if any(kw and kw in subj for kw in subject_keywords):
            final.append(uid)
            logger.info("매칭 메일: UID=%s | Subject=%s", uid.decode(), subj)

    return final


def _fetch_raw_message(imap: imaplib.IMAP4_SSL, uid: bytes) -> bytes:
    """
    FETCH 명령으로 원본 메시지 바이트를 가져옵니다.

    다우오피스 등 일부 서버에서 (RFC822) 형식이 응답 파싱 오류를 내는 경우가
    있어 (BODY.PEEK[])로 fallback합니다.
    """
    for fetch_arg in ("(BODY.PEEK[])", "(RFC822.PEEK)", "(RFC822)"):
        try:
            typ, data = imap.uid("FETCH", uid, fetch_arg)
        except imaplib.IMAP4.error as exc:
            logger.warning("FETCH %s 실패 (UID=%s): %s — 다음 형식 시도",
                           fetch_arg, uid.decode(errors="replace"), exc)
            continue
        if typ != "OK" or not data:
            continue
        # data: list of tuple/bytes pieces
        for part in data:
            if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)) and len(part[1]) > 100:
                return bytes(part[1])
        # 어떤 part도 적합하지 않으면 다음 fetch_arg 시도
    raise RuntimeError(f"메시지 본문을 가져오지 못함 (UID={uid.decode(errors='replace')})")


def download_attachments(
    imap: imaplib.IMAP4_SSL,
    uid: bytes,
    save_dir: Path,
) -> list[Path]:
    """단일 메일 UID에서 .xlsx/.xlsm 첨부파일을 save_dir로 저장."""
    saved: list[Path] = []
    try:
        raw = _fetch_raw_message(imap, uid)
    except Exception as exc:
        logger.error("UID %s 본문 가져오기 실패: %s", uid.decode(errors="replace"), exc)
        return saved

    try:
        msg = email.message_from_bytes(raw)
    except Exception as exc:
        logger.error("UID %s 메시지 파싱 실패: %s", uid.decode(errors="replace"), exc)
        return saved

    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if not filename:
            continue
        try:
            filename = decode_text(filename)
        except Exception:
            pass
        ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
        if ext not in {"xlsx", "xlsm", "xls"}:
            continue

        try:
            payload = part.get_payload(decode=True) or b""
            save_path = save_dir / filename
            save_path.write_bytes(payload)
            saved.append(save_path)
            logger.info("첨부 저장: %s (%d bytes)", save_path.name, len(payload))
        except Exception as exc:
            logger.error("첨부 저장 실패 (%s): %s", filename, exc)

    return saved


# ============================================================
# SMTP: 미수신 알림 메일 발송
# ============================================================
def send_alert_mail(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    to_address: str,
    subject: str,
    body: str,
) -> bool:
    """SSL SMTP로 알림 메일 발송."""
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = to_address
    msg.set_content(body, charset="utf-8")

    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=ctx, timeout=30) as s:
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        logger.info("알림 메일 발송 완료: %s", to_address)
        return True
    except Exception as exc:
        logger.error("알림 메일 발송 실패: %s", exc)
        return False


# ============================================================
# 대시보드 재생성
# ============================================================
def regenerate_dashboard() -> bool:
    """make_dashboard.py를 호출하여 대시보드 HTML 재생성."""
    import subprocess
    here = Path(__file__).parent
    script = here / "make_dashboard.py"
    if not script.exists():
        logger.warning("make_dashboard.py를 찾을 수 없습니다: %s", script)
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(script), "--quiet"],
            cwd=str(here),
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            logger.warning("대시보드 재생성 실패: %s", result.stderr)
            return False
        logger.info("대시보드 재생성 완료")
        return True
    except Exception as exc:
        logger.error("대시보드 재생성 중 예외: %s", exc)
        return False


# ============================================================
# 메인
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=0,
                    help="N일 전부터 오늘까지 검색 (기본 0=오늘만)")
    ap.add_argument("--dry-run", action="store_true",
                    help="실제 업데이트 없이 파싱 결과만 출력")
    ap.add_argument("--no-alert", action="store_true",
                    help="미수신 시 알림 메일 발송 비활성화")
    args = ap.parse_args()

    env = load_env()
    if not env:
        logger.error("환경설정(.env) 로드 실패")
        return 1

    # 필수 키 검증
    required = ["IMAP_HOST", "IMAP_PORT", "EMAIL_ADDRESS", "EMAIL_PASSWORD",
                "ALLOWED_SENDERS", "SUBJECT_KEYWORD",
                "SMTP_HOST", "SMTP_PORT", "ALERT_TO"]
    missing = [k for k in required if not env.get(k)]
    if missing:
        logger.error(".env에 누락된 키: %s", ", ".join(missing))
        return 1

    imap_host = env["IMAP_HOST"]
    imap_port = int(env["IMAP_PORT"])
    email_addr = env["EMAIL_ADDRESS"]
    email_pass = env["EMAIL_PASSWORD"]
    allowed_senders = [a.strip() for a in env["ALLOWED_SENDERS"].split(",") if a.strip()]
    subject_kws = [s.strip() for s in env["SUBJECT_KEYWORD"].split(",") if s.strip()]
    smtp_host = env["SMTP_HOST"]
    smtp_port = int(env["SMTP_PORT"])
    alert_to = env["ALERT_TO"]

    target_dates = []
    today = date.today()
    for i in range(args.days + 1):
        target_dates.append(today - timedelta(days=i))

    # IMAP 접속
    logger.info("IMAP 접속: %s:%d (사용자: %s)", imap_host, imap_port, email_addr)
    try:
        imap = imaplib.IMAP4_SSL(imap_host, imap_port, timeout=30)
        imap.login(email_addr, email_pass)
    except Exception as exc:
        logger.error("IMAP 로그인 실패: %s", exc)
        return 2

    # 임시 다운로드 폴더
    with tempfile.TemporaryDirectory(prefix="작업일보_") as tmpdir:
        tmp_path = Path(tmpdir)
        total_matched = 0
        total_processed = 0

        for d in target_dates:
            uids = search_daily_reports(imap, allowed_senders, subject_kws, d)
            if not uids:
                logger.info("[%s] 매칭되는 작업일보 메일 없음", d.isoformat())
                continue

            total_matched += len(uids)
            for uid in uids:
                try:
                    attachments = download_attachments(imap, uid, tmp_path)
                except imaplib.IMAP4.abort as exc:
                    logger.error("IMAP 세션 abort (UID=%s): %s", uid.decode(errors="replace"), exc)
                    # 세션이 끊겼으면 재연결 후 다음 UID로
                    try:
                        imap.logout()
                    except Exception:
                        pass
                    imap = imaplib.IMAP4_SSL(imap_host, imap_port, timeout=30)
                    imap.login(email_addr, email_pass)
                    imap.select("INBOX", readonly=True)
                    continue
                except Exception as exc:
                    logger.error("첨부 다운로드 예외 (UID=%s): %s", uid.decode(errors="replace"), exc)
                    continue
                if not attachments:
                    logger.warning("UID %s: 엑셀 첨부 없음", uid.decode())
                    continue

                for att in attachments:
                    report = parse_attachment(att)
                    if not report or not report.site_sheet or not report.work_date:
                        logger.warning("파싱 불완전 — 첨부: %s, site=%r, date=%r",
                                       att.name, report and report.site_sheet, report and report.work_date)
                        continue

                    logger.info(
                        "[OK] 파싱 — 현장=%s 일자=%s 인원=%s 작업항목=%d",
                        report.site_sheet, report.work_date,
                        report.total_workers, len(report.items),
                    )

                    if args.dry_run:
                        for it in report.items:
                            logger.info("  · %s | %s | qty=%s",
                                        it.line_normalized, it.work_type, it.quantity)
                        continue

                    # 트래커 업데이트
                    try:
                        result = update_excel(report, EXCEL_PATH, create_backup=True, overwrite=True)
                        if result.errors:
                            for e in result.errors:
                                logger.error("  업데이트 에러: %s", e)
                        if result.skipped:
                            for s in result.skipped:
                                logger.warning("  스킵: %s", s)
                        logger.info("  업데이트 셀: %d개 (백업: %s)",
                                    len(result.results), result.backup_path)
                        total_processed += 1
                    except Exception as exc:
                        logger.error("업데이트 예외: %s", exc)

    imap.logout()

    # 대시보드 재생성
    if total_processed > 0 and not args.dry_run:
        regenerate_dashboard()

    # 미수신 알림
    if total_matched == 0 and not args.no_alert and not args.dry_run:
        body = (
            f"안녕하세요,\n\n"
            f"오늘({today.strftime('%Y-%m-%d')}) 작업일보 메일이 도착하지 않았습니다.\n"
            f"확인 부탁드립니다.\n\n"
            f"- 검색 발신자: {', '.join(allowed_senders)}\n"
            f"- 검색 키워드: {', '.join(subject_kws)}\n"
            f"\n— Auto Tracker"
        )
        send_alert_mail(
            smtp_host, smtp_port, email_addr, email_pass,
            alert_to,
            f"[작업일보 미수신 알림] {today.strftime('%Y-%m-%d')}",
            body,
        )

    logger.info("완료: 매칭 %d건 / 처리 %d건", total_matched, total_processed)
    return 0


if __name__ == "__main__":
    sys.exit(main())
