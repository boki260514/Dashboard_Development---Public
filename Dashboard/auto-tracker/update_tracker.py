"""
카톡 작업일보 → 엑셀 트래킹 시트 자동 입력 메인 스크립트.

사용법:
    1) 클립보드 모드 (기본): 카톡 메시지를 복사한 뒤 실행
        python update_tracker.py

    2) 텍스트 파일 모드: 메시지를 텍스트 파일에 저장한 뒤
        python update_tracker.py message.txt

    3) 직접 입력 모드: 실행 후 메시지를 붙여넣기
        python update_tracker.py --input

옵션:
    --no-backup    백업 생성 안 함 (기본은 생성)
    --dry-run      실제 저장 안 하고 결과만 보기
    --input        터미널에서 직접 입력 받기
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from excel_updater import UpdateReport, update_excel
from parser import DailyReport, parse_message


# ANSI 색상 코드
class Color:
    """터미널 색상 코드."""
    RESET = "\033[0m"
    BOLD = "\033[1m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GRAY = "\033[90m"


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def cprint(message: str, color: str = "", emoji: str = "") -> None:
    """색상과 이모지가 포함된 출력."""
    prefix = f"{emoji} " if emoji else ""
    print(f"{color}{prefix}{message}{Color.RESET}")


def get_message_from_clipboard() -> str:
    """
    클립보드에서 메시지를 읽어옵니다.

    Returns:
        클립보드 텍스트.

    Raises:
        RuntimeError: 클립보드 모듈을 사용할 수 없거나 비어있는 경우.
    """
    try:
        import pyperclip
    except ImportError:
        raise RuntimeError(
            "pyperclip 모듈이 없습니다. "
            "'pip install pyperclip' 명령으로 설치하세요."
        )

    try:
        text = pyperclip.paste()
    except Exception as exc:
        raise RuntimeError(f"클립보드 읽기 실패: {exc}") from exc

    if not text or not text.strip():
        raise RuntimeError(
            "클립보드가 비어있습니다. 카톡 메시지를 먼저 복사하세요."
        )

    return text


def get_message_from_file(filepath: str) -> str:
    """텍스트 파일에서 메시지를 읽어옵니다."""
    return Path(filepath).read_text(encoding="utf-8")


def get_message_from_input() -> str:
    """터미널 직접 입력으로 메시지를 받습니다."""
    cprint("\n카톡 메시지를 붙여넣은 뒤, 빈 줄에서 Ctrl+Z (Windows) 또는 "
           "Ctrl+D (Mac/Linux)를 눌러 종료하세요:\n", Color.YELLOW)
    lines = []
    try:
        while True:
            line = input()
            lines.append(line)
    except EOFError:
        pass
    return "\n".join(lines)


def print_parsed_report(report: DailyReport) -> None:
    """파싱된 보고서를 사람이 읽기 좋게 출력합니다."""
    cprint("\n" + "=" * 60, Color.CYAN)
    cprint("📋 메시지 파싱 결과", Color.CYAN + Color.BOLD)
    cprint("=" * 60, Color.CYAN)

    cprint(f"🏗️  현장: {report.site_raw}", Color.GREEN)
    cprint(f"   → 시트: {report.site_sheet}", Color.GRAY)

    if report.work_date:
        cprint(f"📅 작업일자: {report.work_date.strftime('%Y-%m-%d (%A)')}",
               Color.GREEN)
    else:
        cprint("📅 작업일자: ⚠️  찾을 수 없음", Color.RED)

    if report.work_hours is not None:
        cprint(
            f"⏱️  작업시간: {report.start_time} - {report.end_time} "
            f"({report.work_hours}시간)",
            Color.GREEN,
        )

    if report.total_workers is not None:
        worker_info = f"총원 {report.total_workers}명"
        if report.manager_count is not None and report.worker_count is not None:
            worker_info += (
                f" (관리자 {report.manager_count}, "
                f"작업자 {report.worker_count})"
            )
        cprint(f"👷 작업인원: {worker_info}", Color.GREEN)

    cprint(f"\n📝 작업 항목 ({len(report.items)}개):", Color.GREEN)
    for i, item in enumerate(report.items, 1):
        detail = f" ({item.detail})" if item.detail else ""
        worker_str = f" / 👷 {item.workers}명" if item.workers else ""
        cprint(
            f"   {i}. {item.line_raw} - {item.work_type_raw}: "
            f"{item.quantity}ea{worker_str}{detail}",
            Color.GRAY,
        )

    if report.total_workers and any(it.workers for it in report.items):
        worker_total = sum(it.workers for it in report.items)
        cprint(
            f"\n   👥 인원 분배 합계: {worker_total}명 "
            f"(총원 {report.total_workers}명)",
            Color.CYAN,
        )


def print_update_report(report: UpdateReport) -> None:
    """엑셀 업데이트 결과를 출력합니다."""
    cprint("\n" + "=" * 60, Color.CYAN)
    cprint("💾 엑셀 업데이트 결과", Color.CYAN + Color.BOLD)
    cprint("=" * 60, Color.CYAN)

    if report.backup_path:
        cprint(f"📦 백업 생성: {report.backup_path}", Color.GRAY)

    if report.errors:
        cprint("\n❌ 에러:", Color.RED + Color.BOLD)
        for err in report.errors:
            cprint(f"   • {err}", Color.RED)
        return

    if report.results:
        cprint(f"\n✅ {len(report.results)}개 셀 업데이트:", Color.GREEN)
        for r in report.results:
            mark = "🔄" if r.overwritten else "✨"
            old_str = f"{r.old_value} → " if r.overwritten else ""
            cprint(
                f"   {mark} [{r.cell}] {r.row_label} ({r.field_label}): "
                f"{old_str}{r.new_value}",
                Color.GRAY,
            )

    if report.skipped:
        cprint(f"\n⚠️  건너뜀 ({len(report.skipped)}건):", Color.YELLOW)
        for s in report.skipped:
            cprint(f"   • {s}", Color.YELLOW)
        cprint(
            "\n💡 위 항목을 자동 처리하려면 "
            "config.py의 ROW_MAPPING에 키워드를 추가하세요.",
            Color.GRAY,
        )

    cprint("\n" + "=" * 60, Color.CYAN)


def main() -> int:
    """메인 함수."""
    parser = argparse.ArgumentParser(
        description="카톡 작업일보를 엑셀 트래킹 시트에 자동 입력합니다.",
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="메시지가 담긴 텍스트 파일 (생략 시 클립보드 사용)",
    )
    parser.add_argument(
        "--input",
        action="store_true",
        help="터미널에서 메시지 직접 입력",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="백업 파일 생성 안 함",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="실제로 엑셀에 저장하지 않고 결과만 출력",
    )
    parser.add_argument(
        "--excel",
        help="엑셀 파일 경로 (생략 시 config.py의 EXCEL_PATH 사용)",
    )
    args = parser.parse_args()

    # 메시지 가져오기
    try:
        if args.input:
            message = get_message_from_input()
        elif args.file:
            message = get_message_from_file(args.file)
        else:
            message = get_message_from_clipboard()
            cprint("📋 클립보드에서 메시지 읽음", Color.GREEN)
    except Exception as exc:
        cprint(f"❌ 메시지 로드 실패: {exc}", Color.RED)
        return 1

    if not message.strip():
        cprint("❌ 메시지가 비어있습니다.", Color.RED)
        return 1

    # 파싱
    try:
        report = parse_message(message)
    except ValueError as exc:
        cprint(f"❌ 파싱 실패: {exc}", Color.RED)
        cprint(
            "\n원본 메시지:",
            Color.GRAY,
        )
        cprint(message[:500], Color.GRAY)
        return 1

    print_parsed_report(report)

    # 검증
    if not report.items:
        cprint(
            "\n⚠️  파싱된 작업 항목이 없습니다. "
            "메시지 형식을 확인하세요.",
            Color.YELLOW,
        )
        return 1

    if report.work_date is None:
        cprint(
            "\n❌ 작업 일자를 찾을 수 없어 진행할 수 없습니다.",
            Color.RED,
        )
        return 1

    # Dry run
    if args.dry_run:
        cprint(
            "\n🧪 [DRY RUN 모드] 실제 엑셀 저장은 건너뜁니다.",
            Color.YELLOW,
        )
        return 0

    # 진행 확인
    cprint("\n위 내용을 엑셀에 저장하시겠습니까?", Color.YELLOW + Color.BOLD)
    answer = input("계속하려면 Y, 취소하려면 N 입력: ").strip().lower()
    if answer not in ("y", "yes", ""):
        cprint("❌ 사용자가 취소했습니다.", Color.GRAY)
        return 0

    # 엑셀 업데이트
    excel_path = args.excel  # None이면 config 기본값 사용
    update_kwargs = {"create_backup": not args.no_backup}
    if excel_path:
        update_kwargs["excel_path"] = excel_path

    try:
        update_report = update_excel(report, **update_kwargs)
    except Exception as exc:
        cprint(f"❌ 엑셀 업데이트 중 예외 발생: {exc}", Color.RED)
        logger.exception("Excel update failed")
        return 1

    print_update_report(update_report)

    if update_report.errors:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
