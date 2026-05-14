"""
수동 다운로드한 작업일보 첨부들을 일괄로 HCNSSERVICE_Tracking_v2.xlsx에 반영.

다우오피스 대용량 첨부 등 IMAP으로 자동 수신이 불가능한 경우, 사용자가
.zip을 풀어서 폴더에 넣어두면 이 스크립트가 그 폴더의 모든 .xlsm/.xlsx를
날짜순으로 정렬해 트래커에 누적 반영합니다.

흐름:
    1) 입력 폴더(기본: 오늘_첨부) 하위 모든 .xlsm/.xlsx 수집
    2) excel_parser.parse_attachment로 파싱
    3) (site, date) 둘 다 잡힌 파일만 날짜순 정렬
    4) 첫 파일에만 자동 백업 생성, 나머지는 백업 생략 (한 세션 = 한 백업)
    5) update_excel 호출 (overwrite=True)
    6) 파일별/전체 결과 요약 출력

사용:
    python batch_import_attachments.py                           # 기본 폴더(오늘_첨부)
    python batch_import_attachments.py --dir 다른_폴더            # 다른 폴더
    python batch_import_attachments.py --dry-run                  # 파싱만, 트래커 변경 X
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from excel_parser import parse_attachment
from excel_updater import update_excel
from config import EXCEL_PATH

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default="오늘_첨부",
                    help="작업일보 파일들이 들어있는 폴더 (기본: 오늘_첨부)")
    ap.add_argument("--dry-run", action="store_true",
                    help="파싱만 하고 트래커는 안 건드림")
    args = ap.parse_args()

    src = Path(args.dir)
    if not src.exists():
        print(f"[FAIL] 폴더 없음: {src}", file=sys.stderr)
        return 1

    # 모든 .xlsm/.xlsx (재귀)
    files = sorted(
        list(src.rglob("*.xlsm")) + list(src.rglob("*.xlsx")),
        key=lambda p: p.name,
    )
    if not files:
        print(f"[INFO] 처리할 파일 없음: {src}")
        return 0

    # 1) 파싱
    reports = []  # (path, report)
    skipped = []
    for fp in files:
        r = parse_attachment(fp)
        if r is None:
            skipped.append((fp, "PARSE_FAIL"))
            continue
        if not r.site_sheet:
            skipped.append((fp, f"site_sheet 누락 (project={r.site_raw!r})"))
            continue
        if not r.work_date:
            skipped.append((fp, "work_date 누락"))
            continue
        reports.append((fp, r))

    # 2) 날짜 + 파일명 순 정렬 (같은 날짜 _수정01이 뒤로 가도록 파일명 사전순)
    reports.sort(key=lambda x: (x[1].work_date, x[0].name))

    print("=" * 70)
    print(f"일괄 반영 대상: {len(reports)}개 파일")
    print(f"트래커: {EXCEL_PATH}")
    if args.dry_run:
        print("모드: DRY-RUN (트래커 변경 없음)")
    print("=" * 70)

    total_cells = 0
    total_errs = 0
    total_skips = 0
    backup_path = None

    for idx, (fp, r) in enumerate(reports):
        prefix = f"[{idx+1:>2}/{len(reports)}] {r.site_sheet} {r.work_date}"

        if args.dry_run:
            print(f"{prefix}  items={len(r.items)}  ({fp.name[:50]})")
            continue

        # 첫 번째만 백업 (나머지는 같은 세션이라 의미 없음)
        create_backup = (idx == 0)
        try:
            result = update_excel(r, EXCEL_PATH,
                                  create_backup=create_backup,
                                  overwrite=True)
        except PermissionError:
            print(f"{prefix}  [FAIL] 트래커 잠김. 엑셀을 닫고 다시 실행하세요.")
            return 2
        except Exception as exc:
            print(f"{prefix}  [FAIL] {exc}")
            total_errs += 1
            continue

        if result.backup_path:
            backup_path = result.backup_path

        cells = len(result.results)
        errs = len(result.errors)
        skips = len(result.skipped)
        total_cells += cells
        total_errs += errs
        total_skips += skips

        flag = "OK"
        if errs: flag = "ERR"
        elif skips: flag = "WARN"
        print(f"{prefix}  [{flag}] cells={cells} errs={errs} skips={skips}  ({fp.name[:50]})")
        for e in result.errors[:3]:
            print(f"      err: {e}")
        for s in result.skipped[:3]:
            print(f"      skip: {s}")

    # 3) 요약
    print("=" * 70)
    print(f"파싱 실패/제외: {len(skipped)}개")
    for fp, why in skipped:
        print(f"  - {fp.name}: {why}")
    if not args.dry_run:
        print(f"트래커 업데이트: {total_cells}개 셀")
        print(f"에러: {total_errs}건 / 스킵: {total_skips}건")
        if backup_path:
            print(f"백업: {backup_path}")
    print("=" * 70)
    return 0 if total_errs == 0 else 3


if __name__ == "__main__":
    sys.exit(main())
