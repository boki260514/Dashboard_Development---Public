"""
Mail_Log.xlsx → mail_dashboard.html 생성기.

표시 지표 (스펙 요구):
1. 신규 메일 건수 (최근 7일)
2. 업무 유형별 건수
3. 긴급/높음 우선순위 건수
4. 고객사별 요청 건수 (TOP 10)
5. 금액 합계 (전체 + 최근 7일)
6. 마감일 임박 건수 (7일 이내)
7. 검토필요 건수

사용:
    python mail_dashboard.py            # mail_dashboard.html 생성
    python mail_dashboard.py --quiet
"""
from __future__ import annotations

import argparse
import html
import sys
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

from mail_log_excel import read_all_rows
from mail_log_schema import MAIL_LOG_XLSX, MAIL_DASHBOARD_HTML, WORK_TYPES, PRIORITIES


def _parse_dt(s) -> datetime | None:
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _parse_date(s) -> date | None:
    dt = _parse_dt(s)
    return dt.date() if dt else None


def _num(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def compute_metrics(rows: list[dict]) -> dict:
    """대시보드 표시용 집계 계산."""
    today = date.today()
    week_ago = today - timedelta(days=7)
    week_ahead = today + timedelta(days=7)

    new_last_7d = 0
    work_type_counts: Counter = Counter()
    priority_counts: Counter = Counter()
    high_priority = 0
    company_counts: Counter = Counter()
    total_amount = 0.0
    amount_last_7d = 0.0
    due_soon = 0
    review_needed = 0

    for r in rows:
        recv = _parse_date(r.get("수신일시"))
        if recv and recv >= week_ago:
            new_last_7d += 1
            amount_last_7d += _num(r.get("금액"))

        wt = (r.get("업무유형") or "기타").strip() or "기타"
        work_type_counts[wt] += 1

        pri = (r.get("우선순위") or "보통").strip() or "보통"
        priority_counts[pri] += 1
        if pri in ("긴급", "높음"):
            high_priority += 1

        company = (r.get("회사명") or "").strip()
        if company:
            company_counts[company] += 1

        total_amount += _num(r.get("금액"))

        due = _parse_date(r.get("납기일"))
        if due and today <= due <= week_ahead:
            due_soon += 1

        status = (r.get("상태") or "").strip()
        if status == "검토필요":
            review_needed += 1

    return {
        "total": len(rows),
        "new_last_7d": new_last_7d,
        "work_type_counts": dict(work_type_counts),
        "priority_counts": dict(priority_counts),
        "high_priority": high_priority,
        "top_companies": company_counts.most_common(10),
        "total_amount": total_amount,
        "amount_last_7d": amount_last_7d,
        "due_soon": due_soon,
        "review_needed": review_needed,
    }


def _fmt_money(v: float) -> str:
    return f"{int(v):,}원" if v else "0원"


def render_html(metrics: dict, rows: list[dict]) -> str:
    """HTML 문자열 생성. 외부 의존성 없이 단일 파일로 동작."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 최근 메일 20건
    recent = sorted(
        rows,
        key=lambda r: str(r.get("수신일시") or ""),
        reverse=True,
    )[:20]

    def row_html(r: dict) -> str:
        cells = [
            r.get("수신일시", ""),
            r.get("회사명", ""),
            r.get("업무유형", ""),
            r.get("제목", ""),
            r.get("우선순위", ""),
            r.get("상태", ""),
            f"{int(_num(r.get('금액'))):,}" if _num(r.get("금액")) else "",
            r.get("납기일", ""),
            r.get("담당자", ""),
        ]
        tds = "".join(f"<td>{html.escape(str(c or ''))}</td>" for c in cells)
        # 상태에 따라 행 색상
        klass = ""
        st = (r.get("상태") or "").strip()
        if st == "검토필요":
            klass = ' class="row-review"'
        elif (r.get("우선순위") or "") == "긴급":
            klass = ' class="row-urgent"'
        return f"<tr{klass}>{tds}</tr>"

    work_type_rows = "".join(
        f'<li><span>{html.escape(wt)}</span><b>{metrics["work_type_counts"].get(wt, 0)}</b></li>'
        for wt in WORK_TYPES
    )
    priority_rows = "".join(
        f'<li><span>{html.escape(p)}</span><b>{metrics["priority_counts"].get(p, 0)}</b></li>'
        for p in PRIORITIES
    )
    company_rows = "".join(
        f'<li><span>{html.escape(c)}</span><b>{n}</b></li>'
        for c, n in metrics["top_companies"]
    ) or '<li class="empty">데이터 없음</li>'

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>메일 로그 대시보드</title>
<style>
  body {{ font-family: 'Segoe UI', 'Apple SD Gothic Neo', sans-serif; margin:0; padding:24px; background:#f4f6fa; color:#1d2433; }}
  h1 {{ margin:0 0 4px; }}
  .meta {{ color:#6b7280; font-size:13px; margin-bottom:24px; }}
  .grid {{ display:grid; grid-template-columns:repeat(4, 1fr); gap:16px; margin-bottom:24px; }}
  .card {{ background:white; border-radius:10px; padding:16px 18px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
  .card .label {{ font-size:12px; color:#6b7280; margin-bottom:6px; }}
  .card .value {{ font-size:28px; font-weight:700; color:#1f2937; }}
  .card.urgent .value {{ color:#dc2626; }}
  .card.review .value {{ color:#d97706; }}
  .card.money .value {{ font-size:22px; }}
  .panels {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:16px; margin-bottom:24px; }}
  .panel {{ background:white; border-radius:10px; padding:14px 18px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
  .panel h3 {{ margin:0 0 10px; font-size:14px; color:#374151; }}
  .panel ul {{ list-style:none; padding:0; margin:0; }}
  .panel li {{ display:flex; justify-content:space-between; padding:5px 0; border-bottom:1px solid #f3f4f6; font-size:13px; }}
  .panel li:last-child {{ border-bottom:none; }}
  .panel li.empty {{ color:#9ca3af; font-style:italic; }}
  table {{ width:100%; border-collapse:collapse; background:white; border-radius:10px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
  th, td {{ padding:9px 12px; text-align:left; font-size:13px; border-bottom:1px solid #f3f4f6; }}
  th {{ background:#305496; color:white; font-weight:600; }}
  tr.row-review td {{ background:#fff7ed; }}
  tr.row-urgent td {{ background:#fef2f2; }}
  .section-title {{ font-size:14px; font-weight:600; color:#374151; margin:0 0 10px; }}
</style>
</head>
<body>
  <h1>📧 메일 로그 대시보드</h1>
  <div class="meta">최종 갱신: {now} · 누적 {metrics['total']:,}건</div>

  <div class="grid">
    <div class="card"><div class="label">최근 7일 신규</div><div class="value">{metrics['new_last_7d']:,}건</div></div>
    <div class="card urgent"><div class="label">긴급+높음 우선순위</div><div class="value">{metrics['high_priority']:,}건</div></div>
    <div class="card review"><div class="label">검토필요</div><div class="value">{metrics['review_needed']:,}건</div></div>
    <div class="card urgent"><div class="label">마감일 임박(≤7일)</div><div class="value">{metrics['due_soon']:,}건</div></div>
    <div class="card money"><div class="label">금액 합계 (전체)</div><div class="value">{_fmt_money(metrics['total_amount'])}</div></div>
    <div class="card money"><div class="label">금액 합계 (최근 7일)</div><div class="value">{_fmt_money(metrics['amount_last_7d'])}</div></div>
  </div>

  <div class="panels">
    <div class="panel">
      <h3>업무 유형별</h3>
      <ul>{work_type_rows}</ul>
    </div>
    <div class="panel">
      <h3>우선순위별</h3>
      <ul>{priority_rows}</ul>
    </div>
    <div class="panel">
      <h3>고객사 TOP 10</h3>
      <ul>{company_rows}</ul>
    </div>
  </div>

  <div class="section-title">최근 메일 20건</div>
  <table>
    <thead><tr>
      <th>수신일시</th><th>회사명</th><th>업무유형</th><th>제목</th>
      <th>우선순위</th><th>상태</th><th>금액</th><th>납기일</th><th>담당자</th>
    </tr></thead>
    <tbody>{"".join(row_html(r) for r in recent) or '<tr><td colspan="9" style="text-align:center;color:#9ca3af;padding:20px">데이터 없음</td></tr>'}</tbody>
  </table>
</body>
</html>"""


def regenerate(
    excel_path: Path = MAIL_LOG_XLSX,
    output_path: Path = MAIL_DASHBOARD_HTML,
    quiet: bool = False,
) -> bool:
    """Mail_Log.xlsx 읽어 mail_dashboard.html 재생성. 성공 시 True."""
    rows = read_all_rows(excel_path)
    metrics = compute_metrics(rows)
    html_str = render_html(metrics, rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_text(html_str, encoding="utf-8")
    if not quiet:
        print(f"[OK] {output_path} 생성됨 (누적 {metrics['total']}건)")
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    try:
        regenerate(quiet=args.quiet)
    except Exception as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
