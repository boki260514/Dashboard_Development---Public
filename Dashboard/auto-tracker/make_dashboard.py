"""
대시보드 자동 갱신 메인 스크립트.

각 현장별로 별도 HTML 파일을 생성하고, 메인 dashboard.html은
가장 활발한 현장(또는 config의 DEFAULT_ACTIVE_SITE)을 보여줍니다.

생성되는 파일:
    dashboard.html                  ← 메인 (가장 활발한 현장)
    dashboard_CHW1FC.html
    dashboard_CHW2FC.html
    dashboard_ECH2FC.html
    dashboard_SIH2FC.html
    dashboard_DON1FC.html
    dashboard_MCN1FC.html
    dashboard_etc.html              ← None PJT

사용법:
    python make_dashboard.py
"""
from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

from config import EXCEL_PATH, OUTPUT_PATH, DEFAULT_ACTIVE_SITE
from extract_data import extract_dashboard_data
from render_html import render_dashboard, site_to_filename


# ANSI 색상
class C:
    R = "\033[0m"
    BOLD = "\033[1m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    CYAN = "\033[96m"
    GRAY = "\033[90m"


def main() -> int:
    p = argparse.ArgumentParser(description="트래킹 시트 → 대시보드 HTML 자동 생성")
    p.add_argument("--excel", help="엑셀 파일 경로 (기본: config.py의 EXCEL_PATH)")
    p.add_argument("--output", help="메인 출력 HTML 경로 (기본: config.py의 OUTPUT_PATH)")
    p.add_argument("--site", help="메인에 표시할 현장 코드 (예: CHW1FC)")
    p.add_argument("--quiet", action="store_true", help="자세한 출력 안 함")
    args = p.parse_args()

    excel_path = args.excel or EXCEL_PATH
    output_path = args.output or OUTPUT_PATH
    active_site = args.site or DEFAULT_ACTIVE_SITE

    if not args.quiet:
        print(f"{C.CYAN}{'=' * 60}{C.R}")
        print(f"{C.BOLD}📊 대시보드 자동 갱신{C.R}")
        print(f"{C.CYAN}{'=' * 60}{C.R}")
        print(f"📂 엑셀: {C.GRAY}{excel_path}{C.R}")
        print(f"📄 출력: {C.GRAY}{output_path}{C.R}")
        if active_site:
            print(f"🏗️ 메인 현장: {C.GRAY}{active_site}{C.R}")
        print()

    # 엑셀 파일 존재 확인
    if not Path(excel_path).exists():
        print(f"{C.RED}❌ 엑셀 파일을 찾을 수 없습니다: {excel_path}{C.R}")
        print(f"   config.py의 EXCEL_PATH를 확인하세요.")
        return 1

    # 데이터 추출
    try:
        if not args.quiet:
            print(f"{C.YELLOW}⏳ 엑셀 데이터 읽는 중...{C.R}")
        data = extract_dashboard_data(excel_path)
        if not args.quiet:
            print(f"{C.GREEN}✅ 데이터 추출 완료{C.R}")
            print(f"   현장 {len(data.sites)}개")
            print(f"   최근 활동 {len(data.recent_activity)}건")
            print(f"   일별 추이 {len(data.daily_trend)}일")
    except Exception as exc:
        print(f"{C.RED}❌ 엑셀 읽기 실패: {exc}{C.R}")
        traceback.print_exc()
        return 1

    # 현장별 요약 출력
    if not args.quiet:
        print(f"\n{C.BOLD}📋 현장별 진행률:{C.R}")
        for site in data.sites:
            pct = site.overall_pct * 100
            mark = "🟢" if pct >= 50 else ("🟡" if pct >= 10 else "🔴")
            display = "기타" if site.site_code == "None PJT" else site.site_code
            print(
                f"   {mark} {display:<10} "
                f"{site.total_actual:>6.0f} / {site.total_plan:>6.0f}  "
                f"({pct:>5.1f}%)  · {len(site.items)}개 항목"
            )

    # 메인에 표시할 현장 결정
    if active_site is None:
        non_empty = [s for s in data.sites if s.total_actual > 0]
        if non_empty:
            main_site_code = max(non_empty, key=lambda s: s.total_actual).site_code
        else:
            main_site_code = data.sites[0].site_code if data.sites else None
    else:
        main_site_code = active_site

    # 출력 디렉토리
    output_dir = Path(output_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1) 메인 dashboard.html (요청된 또는 자동 선택된 현장)
    try:
        if not args.quiet:
            print(f"\n{C.YELLOW}⏳ HTML 생성 중...{C.R}")

        html = render_dashboard(data, active_site_code=main_site_code)
        Path(output_path).write_text(html, encoding="utf-8")
        if not args.quiet:
            display = "기타" if main_site_code == "None PJT" else main_site_code
            print(f"   ✅ 메인 (dashboard.html) — {display}")

        # 2) 각 현장별 HTML 생성
        for site in data.sites:
            site_html = render_dashboard(data, active_site_code=site.site_code)
            site_filename = site_to_filename(site.site_code)
            site_path = output_dir / site_filename
            site_path.write_text(site_html, encoding="utf-8")
            if not args.quiet:
                display = "기타" if site.site_code == "None PJT" else site.site_code
                print(f"   ✅ {site_filename} — {display}")

    except Exception as exc:
        print(f"{C.RED}❌ HTML 생성 실패: {exc}{C.R}")
        traceback.print_exc()
        return 1

    if not args.quiet:
        print(f"\n{C.GREEN}{C.BOLD}🎉 대시보드 생성 완료!{C.R}")
        print(f"   📄 파일: {C.CYAN}{Path(output_path).resolve()}{C.R}")
        print(f"   📊 현장 페이지 {len(data.sites)}개 추가 생성됨")
        print(f"   ⏰ 생성 시각: {data.generated_at.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"\n{C.GRAY}💡 dashboard.html을 더블클릭하면 브라우저에서 열립니다.{C.R}")
        print(f"{C.GRAY}   상단 탭을 클릭하면 다른 현장으로 이동합니다.{C.R}")
        print(f"{C.CYAN}{'=' * 60}{C.R}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
