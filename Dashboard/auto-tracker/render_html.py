"""
DashboardData를 HTML 대시보드로 렌더링하는 모듈.

후보 1 (v3 노션 스타일, 하얀 배경) 디자인을 기반으로 하며,
extract_data.py에서 추출한 실제 데이터를 채워 넣습니다.
"""
from __future__ import annotations

import json
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from extract_data import DashboardData, SiteData, WorkItem


# ============================================================
# 상태 판정 로직
# ============================================================

def get_item_status(item: WorkItem) -> str:
    """
    작업 항목의 상태 판정.
    
    - good: 진행률 정상 (50%+ 또는 적당히 진행 중)
    - warn: 진행 중이지만 페이스 느림 (1~30% 사이)
    - bad: 거의 시작 안 함 또는 매우 느림 (<5% & 일정 경과)
    - done: 100% 완료
    """
    if item.plan_total <= 0:
        return "warn"  # 계획 없음
    pct = item.progress_pct
    if pct >= 1.0:
        return "done"
    if pct >= 0.4:
        return "good"
    if pct >= 0.1:
        return "warn"
    return "bad"


def status_to_label(status: str) -> tuple[str, str]:
    """상태를 한글 라벨 + 이모지로 변환."""
    return {
        "good": ("진행 중", "🟢"),
        "warn": ("점검 필요", "🟡"),
        "bad":  ("지연", "🔴"),
        "done": ("완료", "✅"),
    }.get(status, ("진행 중", "🟢"))


def status_class(status: str) -> str:
    """상태를 CSS 클래스로 변환."""
    return {
        "good": "good",
        "warn": "warn",
        "bad": "bad",
        "done": "good",
    }.get(status, "good")


def work_type_emoji(work_type: str) -> str:
    """작업 종류별 이모지."""
    return {
        "Gap guarding": "🔨",
        "Cover": "🛡️",
        "Cable": "🔌",
        "Gap Plate": "📐",
    }.get(work_type, "🔧")


def work_type_tag_class(work_type: str) -> str:
    """작업 종류별 태그 컬러 클래스."""
    return {
        "Gap guarding": "sage",
        "Cover": "mocha",
        "Cable": "terra",
        "Gap Plate": "plum",
    }.get(work_type, "slate")


def line_tag_class(short_name: str) -> str:
    """라인 이름별 태그 컬러."""
    if "PB" in short_name:
        return "mocha"
    if "CB" in short_name:
        return "plum"
    if "ST" in short_name:
        return "slate"
    return "slate"


# ============================================================
# HTML 부분 생성기
# ============================================================

def _fmt_mday(v: float) -> str:
    """멘데이 값을 표시용 문자열로. 0이거나 음수면 '-' 반환."""
    if v is None or v <= 0:
        return "-"
    # 정수면 소수점 없이, 아니면 1자리
    if abs(v - round(v)) < 0.05:
        return f"{int(round(v))}"
    return f"{v:.1f}"


def render_kpi_section(site: SiteData, data: DashboardData) -> str:
    """상단 KPI 카드 HTML. (전체 진행률 + 작업 일수 + 총 계획 M-day + 누적 사용 M-day + 금일 사용 M-day)"""
    pct = site.overall_pct * 100
    pct_int = int(round(pct))

    # 진행률 링 계산 (반지름 34, 둘레 213.6)
    circumference = 213.6
    offset = circumference * (1 - site.overall_pct)

    # 멘데이 표시값
    plan_md = _fmt_mday(site.total_plan_mday)
    used_md = _fmt_mday(site.total_used_mday)
    today_md = _fmt_mday(site.total_today_mday)

    # 멘데이 사용률 (계획 대비 누적 사용)
    if site.total_plan_mday > 0:
        mday_pct = (site.total_used_mday / site.total_plan_mday) * 100
        mday_pct_str = f"{mday_pct:.0f}%"
        mday_pct_bar = min(mday_pct, 100)
        plan_remaining = max(site.total_plan_mday - site.total_used_mday, 0)
        plan_sub = f"잔여 {_fmt_mday(plan_remaining)} md"
    else:
        mday_pct_str = "—"
        mday_pct_bar = 0
        plan_sub = "config.py에 입력"

    return f"""
<div class="hero-stats">
  <div class="stat-card featured">
    <div class="sc-label">📊 전체 진행률</div>
    <div class="sc-value">{site.total_actual:.0f}<span class="unit">/ {site.total_plan:.0f} sets</span></div>
    <div class="ring-vis">
      <svg width="80" height="80">
        <circle class="ring-bg" cx="40" cy="40" r="34"></circle>
        <circle class="ring-fill" cx="40" cy="40" r="34"
          stroke-dasharray="{circumference}" stroke-dashoffset="{offset:.1f}"></circle>
      </svg>
      <div style="position:absolute; inset:0; display:flex; align-items:center; justify-content:center; font-family:'Pretendard Variable', sans-serif; font-weight:200; font-size:22px; letter-spacing:-0.03em;">{pct_int}%</div>
    </div>
    <div class="sc-meta">
      <span>완료까지 {site.total_plan - site.total_actual:.0f} sets</span>
      <span class="sc-trend up">진행 중</span>
    </div>
    <div class="sc-progress"><div style="width:{pct:.1f}%"></div></div>
  </div>

  <div class="stat-card">
    <div class="sc-label">📅 작업 일수</div>
    <div class="sc-value">{site.days_elapsed}<span class="unit">일</span></div>
    <div class="sc-meta">
      <span>실제 작업 진행일</span>
      <span class="sc-trend up">기록됨</span>
    </div>
    <div class="sc-progress"><div style="width:{min(site.days_elapsed*2, 100)}%; background: linear-gradient(90deg, var(--sage), #6b8e58);"></div></div>
  </div>

  <div class="stat-card">
    <div class="sc-label">📋 총 계획 M-day</div>
    <div class="sc-value">{plan_md}<span class="unit">md</span></div>
    <div class="sc-meta">
      <span>{plan_sub}</span>
      <span class="sc-trend warn">계획</span>
    </div>
    <div class="sc-progress"><div style="width:100%; background: linear-gradient(90deg, var(--slate-bg), var(--slate));"></div></div>
  </div>

  <div class="stat-card">
    <div class="sc-label">📈 누적 사용 M-day</div>
    <div class="sc-value">{used_md}<span class="unit">md</span></div>
    <div class="sc-meta">
      <span>사용률 {mday_pct_str}</span>
      <span class="sc-trend up">누적</span>
    </div>
    <div class="sc-progress"><div style="width:{mday_pct_bar:.1f}%; background: linear-gradient(90deg, var(--plum), #4a3450);"></div></div>
  </div>

  <div class="stat-card">
    <div class="sc-label">⏱️ 금일 사용 M-day</div>
    <div class="sc-value">{today_md}<span class="unit">md</span></div>
    <div class="sc-meta">
      <span>오늘 투입 인원</span>
      <span class="sc-trend up">오늘</span>
    </div>
    <div class="sc-progress"><div style="width:{min((site.total_today_mday*10), 100):.0f}%; background: linear-gradient(90deg, var(--gold-bg), var(--gold));"></div></div>
  </div>
</div>
"""



def render_work_card(item: WorkItem) -> str:
    """공종별 카드 1개 HTML."""
    status = get_item_status(item)
    label, emoji = status_to_label(status)
    cls = status_class(status)
    
    # 태그 빌드
    tags = [
        f'<span class="tag {line_tag_class(item.short_name)}">{item.short_name}</span>',
        f'<span class="tag {work_type_tag_class(item.work_type)}">{item.work_type}</span>',
    ]
    if item.line_position:
        tags.append(f'<span class="tag slate">{item.line_position}</span>')
    
    # 단위 결정 (sets vs ea)
    unit = "sets" if "guarding" in item.work_type.lower() or "cover" in item.work_type.lower() else "ea"
    
    pct = item.progress_pct * 100
    
    # 진행률이 0이면 "대기 중" 표시
    if item.plan_total <= 0:
        plan_display = "계획 미설정"
    else:
        plan_display = f"{item.actual_total:.0f} / {item.plan_total:.0f} {unit}"
    
    return f"""
  <div class="wcard">
    <div class="wcard-status {cls}"></div>
    <div class="wcard-head">
      <div class="wcard-icon {cls}">{work_type_emoji(item.work_type)}</div>
      <div class="wcard-status-badge {cls}"><span class="dot"></span>{label}</div>
    </div>
    <div class="wcard-title">{item.short_name} {item.work_type}</div>
    <div class="wcard-subtitle">{item.label}</div>
    <div class="wcard-tags">
      {''.join(tags)}
    </div>
    <div class="wcard-progress">
      <div class="wcard-pp">
        <span class="wcard-pp-label">{plan_display}</span>
        <span class="wcard-pp-value">{pct:.1f}<span class="unit">%</span></span>
      </div>
      <div class="wcard-bar"><div class="{cls}" style="width:{min(pct, 100):.1f}%"></div></div>
    </div>
  </div>"""


def render_work_section(site: SiteData) -> str:
    """공종별 카드 섹션 HTML."""
    if not site.items:
        return '<div style="padding:40px; text-align:center; color:var(--soft);">작업 항목 없음</div>'
    
    # 진행률 높은 순으로 정렬 (활성 항목 우선)
    items_sorted = sorted(
        site.items,
        key=lambda it: (it.actual_total > 0, it.progress_pct),
        reverse=True,
    )
    
    cards_html = "\n".join(render_work_card(it) for it in items_sorted)
    return f'<div class="work-grid">{cards_html}</div>'


def render_activity_item(record: dict) -> str:
    """활동 로그 한 항목 HTML."""
    d = record["date"]
    weekday = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"][d.weekday()]
    
    # 가장 최근(오늘 또는 가까운 날)이면 Active, 아니면 Done
    days_ago = (date.today() - d).days
    if days_ago <= 1:
        status_tag = '<span class="tag sage" style="font-size:10px;">Active</span>'
    else:
        status_tag = '<span class="tag slate" style="font-size:10px;">Done</span>'
    
    # 상세 정보
    details = []
    if record["workers"]:
        details.append(f"{record['workers']}명")
    if record["install_hour"]:
        details.append(f"{record['install_hour']:.1f}/h")
    detail_str = " · ".join(details) if details else "기록 완료"
    
    unit = "sets" if "guarding" in record["work_type"].lower() else "ea"
    
    return f"""
      <div class="activity-item">
        <div class="activity-date">{d.day:02d}<small>{weekday}</small></div>
        <div class="activity-content">
          <div class="title">{record['short_name']} {record['work_type']} {status_tag}</div>
          <div class="meta">{record['site']} · {detail_str}</div>
        </div>
        <div class="activity-qty">{record['qty']:.0f}<small>{unit}</small></div>
      </div>"""


def render_activity_section(records: list[dict]) -> str:
    """활동 로그 섹션 HTML."""
    if not records:
        return '<div style="padding:40px; text-align:center; color:var(--soft);">최근 작업 기록 없음</div>'
    items_html = "\n".join(render_activity_item(r) for r in records[:8])
    return f'<div class="activity">{items_html}</div>'


def render_chart_data(daily_trend: list[tuple[date, float]]) -> tuple[str, str]:
    """차트용 라벨/데이터 JSON 문자열 반환."""
    labels = [f"{d.month}/{d.day}" for d, _ in daily_trend]
    data = [v for _, v in daily_trend]
    return json.dumps(labels, ensure_ascii=False), json.dumps(data)


def site_to_filename(site_code: str) -> str:
    """
    시트 코드를 안전한 HTML 파일명으로 변환.
    
    "CHW1FC"      → "dashboard_CHW1FC.html"
    "None PJT"    → "dashboard_etc.html"
    """
    if site_code == "None PJT":
        return "dashboard_etc.html"
    # 공백, 특수문자 제거
    safe = "".join(c if c.isalnum() else "_" for c in site_code)
    return f"dashboard_{safe}.html"


def render_tabs(sites: list[SiteData], active_code: str) -> str:
    """현장 탭 HTML — 클릭 가능한 링크로."""
    tabs = []
    for site in sites:
        is_active = " active" if site.site_code == active_code else ""
        active_count = sum(1 for it in site.items if it.actual_total > 0)
        display_name = "기타" if site.site_code == "None PJT" else site.site_code
        href = site_to_filename(site.site_code)
        tabs.append(
            f'<a class="tab{is_active}" href="{href}">'
            f'🏗️ {display_name} '
            f'<span class="count">{active_count}</span></a>'
        )
    return f'<div class="tabs-wrap">{"".join(tabs)}</div>'


# ============================================================
# 전체 HTML 생성
# ============================================================

# HTML 템플릿 (후보 1 v3 - 하얀 배경)
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>프로젝트 진척 대시보드 · {site_code}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
  :root {{
    --bg: #ffffff;
    --bg-2: #f5f3ee;
    --paper: #ffffff;
    --paper-warm: #fafaf8;
    --ink: #1f1d1a;
    --ink-2: #3a3631;
    --soft: #6b665d;
    --softer: #8e887e;
    --line: #ececec;
    --line-strong: #d8d8d8;

    --mocha: #6f4e37;
    --mocha-bg: #ede2d3;
    --mocha-light: #b08968;
    --terracotta: #c2410c;
    --terracotta-bg: #fae5d3;
    --sage: #87976e;
    --sage-bg: #dde4d0;
    --olive: #6b7d3a;
    --plum: #6b4e71;
    --plum-bg: #e8dcec;
    --slate: #475569;
    --slate-bg: #dbe2ea;
    --gold: #b8860b;
    --gold-bg: #f5ebd0;

    --shadow-soft: 0 1px 2px rgba(31,29,26,0.04), 0 4px 12px rgba(31,29,26,0.04);
    --shadow-hover: 0 2px 4px rgba(31,29,26,0.06), 0 12px 32px rgba(31,29,26,0.08);
  }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  html, body {{
    font-family: 'Pretendard Variable', 'Pretendard', -apple-system, sans-serif;
    background: var(--bg);
    color: var(--ink);
    line-height: 1.5;
    -webkit-font-smoothing: antialiased;
    font-feature-settings: "ss01", "ss02";
  }}
  .serif {{ font-family: 'Pretendard Variable', 'Pretendard', sans-serif; font-weight: 200; }}
  .mono  {{ font-family: 'JetBrains Mono', monospace; }}

  body::before {{
    content: '';
    position: fixed; inset: 0;
    background-image:
      radial-gradient(circle at 20% 10%, rgba(184,134,11,0.025) 0, transparent 40%),
      radial-gradient(circle at 80% 80%, rgba(135,151,110,0.025) 0, transparent 40%);
    pointer-events: none; z-index: 0;
  }}
  .page {{ max-width: 1240px; margin: 0 auto; padding: 48px 32px 96px; position: relative; z-index: 1; }}

  header {{ margin-bottom: 40px; display: grid; grid-template-columns: 1fr auto; align-items: end; gap: 32px; }}
  .breadcrumb {{ font-size: 13px; color: var(--soft); margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }}
  .breadcrumb a {{ color: var(--soft); text-decoration: none; transition: color 0.15s; }}
  .breadcrumb a:hover {{ color: var(--ink); }}
  .breadcrumb .sep {{ color: var(--line-strong); }}
  .breadcrumb .current {{ color: var(--ink); font-weight: 500; }}

  h1 {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-weight: 200;
    font-size: 64px;
    letter-spacing: -0.03em;
    line-height: 1.05;
    margin-bottom: 20px;
    color: var(--ink);
  }}
  h1 .label {{ font-weight: 700; letter-spacing: -0.02em; }}
  h1 .accent {{ color: var(--terracotta); font-weight: 200; }}

  .meta-pills {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  .meta-pill {{
    display: inline-flex; align-items: center; gap: 6px;
    padding: 6px 12px;
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 100px;
    font-size: 12px;
    color: var(--ink-2);
  }}
  .meta-pill .dot {{ width: 6px; height: 6px; border-radius: 50%; background: var(--sage); }}
  .meta-pill.live .dot {{
    background: #22c55e;
    animation: livePulse 2s ease-in-out infinite;
    box-shadow: 0 0 0 0 rgba(34,197,94,0.6);
  }}
  @keyframes livePulse {{
    0%, 100% {{ box-shadow: 0 0 0 0 rgba(34,197,94,0.6); }}
    50% {{ box-shadow: 0 0 0 6px rgba(34,197,94,0); }}
  }}

  .refresh-btn {{
    margin-left: auto;
    display: inline-flex; align-items: center; gap: 8px;
    padding: 10px 16px;
    background: var(--ink);
    color: var(--bg);
    border: none;
    border-radius: 10px;
    font-family: inherit;
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.15s;
  }}
  .refresh-btn:hover {{ background: var(--mocha); transform: translateY(-1px); }}

  .tabs-wrap {{
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 6px;
    display: flex;
    gap: 2px;
    margin-bottom: 32px;
    overflow-x: auto;
    box-shadow: var(--shadow-soft);
  }}
  .tab {{
    padding: 10px 18px;
    font-size: 13px;
    font-weight: 500;
    color: var(--soft);
    border-radius: 10px;
    cursor: pointer;
    transition: all 0.15s;
    white-space: nowrap;
    display: flex; align-items: center; gap: 8px;
    text-decoration: none;
  }}
  .tab:hover {{ background: var(--bg-2); color: var(--ink); }}
  .tab.active {{ background: var(--ink); color: var(--bg); box-shadow: 0 2px 6px rgba(31,29,26,0.15); }}
  .tab .count {{ font-size: 11px; padding: 2px 7px; border-radius: 100px; background: rgba(0,0,0,0.06); }}
  .tab.active .count {{ background: rgba(255,255,255,0.15); color: var(--bg); }}

  .hero-stats {{ display: grid; grid-template-columns: 1.4fr 1fr 1fr 1fr 1fr; gap: 14px; margin-bottom: 36px; }}
  .stat-card {{
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 22px 20px;
    position: relative;
    transition: all 0.2s;
    overflow: hidden;
  }}
  .stat-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-hover); }}
  .stat-card.featured {{
    background: linear-gradient(135deg, var(--ink) 0%, #2d2823 100%);
    color: var(--bg);
    border-color: var(--ink);
  }}
  .stat-card.featured::before {{
    content: '';
    position: absolute;
    top: -50%; right: -20%;
    width: 280px; height: 280px;
    background: radial-gradient(circle, rgba(194,65,12,0.25) 0%, transparent 60%);
    pointer-events: none;
  }}
  .stat-card.featured::after {{
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, var(--terracotta), var(--gold));
  }}

  .sc-label {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: var(--softer);
    margin-bottom: 14px;
    font-weight: 500;
  }}
  .stat-card.featured .sc-label {{ color: rgba(246,243,237,0.55); }}

  .sc-value {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-weight: 200;
    font-size: 48px;
    line-height: 0.95;
    letter-spacing: -0.04em;
    color: var(--ink);
    display: flex;
    align-items: baseline;
    gap: 6px;
  }}
  .stat-card.featured .sc-value {{ color: var(--bg); }}
  .sc-value .unit {{
    font-family: 'Pretendard Variable', sans-serif;
    font-size: 18px;
    color: var(--soft);
    font-weight: 400;
  }}
  .stat-card.featured .sc-value .unit {{ color: rgba(246,243,237,0.6); }}

  .sc-meta {{
    margin-top: 16px;
    display: flex; justify-content: space-between; align-items: center;
    font-size: 12px;
    color: var(--soft);
  }}
  .stat-card.featured .sc-meta {{ color: rgba(246,243,237,0.7); }}
  .sc-trend {{
    display: inline-flex; align-items: center; gap: 4px;
    padding: 3px 8px;
    border-radius: 100px;
    font-weight: 500;
  }}
  .sc-trend.up {{ color: #15803d; background: #dcfce7; }}
  .sc-trend.down {{ color: #b91c1c; background: #fee2e2; }}
  .sc-trend.warn {{ color: #a16207; background: #fef3c7; }}
  .stat-card.featured .sc-trend.up {{ background: rgba(34,197,94,0.2); color: #86efac; }}

  .sc-progress {{
    margin-top: 14px;
    height: 6px;
    background: var(--bg-2);
    border-radius: 100px;
    overflow: hidden;
  }}
  .stat-card.featured .sc-progress {{ background: rgba(255,255,255,0.1); }}
  .sc-progress > div {{
    height: 100%;
    border-radius: 100px;
    background: linear-gradient(90deg, var(--mocha-light), var(--mocha));
    transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
  }}
  .stat-card.featured .sc-progress > div {{
    background: linear-gradient(90deg, var(--terracotta), var(--gold));
  }}

  .ring-vis {{ position: absolute; top: 22px; right: 20px; width: 72px; height: 72px; }}
  .ring-vis svg {{ transform: rotate(-90deg); }}
  .ring-bg {{ fill: none; stroke: rgba(255,255,255,0.1); stroke-width: 6; }}
  .ring-fill {{ fill: none; stroke: var(--terracotta); stroke-width: 6; stroke-linecap: round; }}

  /* 멘데이(M-day) 미니 스탯 (전체 진행률 카드 내부) */
  .mday-strip {{
    margin-top: 18px;
    padding-top: 16px;
    border-top: 1px dashed rgba(255,255,255,0.15);
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 12px;
    position: relative;
    z-index: 1;
  }}
  .mday-item {{
    padding: 10px 12px;
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 10px;
    transition: background 0.2s;
  }}
  .mday-item:hover {{ background: rgba(255,255,255,0.10); }}
  .mday-label {{
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: rgba(246,243,237,0.6);
    margin-bottom: 6px;
    font-weight: 500;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .mday-value {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-size: 26px;
    font-weight: 200;
    letter-spacing: -0.03em;
    line-height: 1;
    color: var(--bg);
    display: flex;
    align-items: baseline;
    gap: 4px;
  }}
  .mday-unit {{
    font-size: 12px;
    color: rgba(246,243,237,0.55);
    font-weight: 400;
  }}
  .mday-sub {{
    margin-top: 6px;
    font-size: 11px;
    color: rgba(246,243,237,0.55);
  }}

  .section-head {{ display: flex; align-items: baseline; justify-content: space-between; margin-bottom: 20px; padding-bottom: 12px; }}
  .section-title {{ font-size: 22px; font-weight: 700; letter-spacing: -0.01em; display: flex; align-items: center; gap: 12px; }}
  .section-title .num {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-style: italic;
    font-weight: 300;
    font-size: 16px;
    color: var(--terracotta);
    border: 1.5px solid var(--terracotta);
    width: 28px; height: 28px;
    display: inline-flex; align-items: center; justify-content: center;
    border-radius: 50%;
    background: var(--terracotta-bg);
  }}
  .section-action {{
    font-size: 13px; color: var(--soft); cursor: pointer;
    display: inline-flex; align-items: center; gap: 4px;
    padding: 4px 10px; border-radius: 6px;
  }}
  .section-action:hover {{ background: var(--paper); color: var(--ink); }}

  .work-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; margin-bottom: 40px; }}
  .wcard {{
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 24px;
    cursor: pointer;
    transition: all 0.2s;
    display: flex; flex-direction: column;
    position: relative;
    overflow: hidden;
  }}
  .wcard:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-hover); border-color: var(--line-strong); }}
  .wcard-status {{ position: absolute; top: 0; left: 0; right: 0; height: 3px; }}
  .wcard-status.good {{ background: linear-gradient(90deg, var(--sage), #6b8e58); }}
  .wcard-status.warn {{ background: linear-gradient(90deg, var(--gold), #d97706); }}
  .wcard-status.bad  {{ background: linear-gradient(90deg, var(--terracotta), #9a2e08); }}

  .wcard-head {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; }}
  .wcard-icon {{ width: 40px; height: 40px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 20px; }}
  .wcard-icon.good {{ background: var(--sage-bg); }}
  .wcard-icon.warn {{ background: var(--gold-bg); }}
  .wcard-icon.bad  {{ background: var(--terracotta-bg); }}

  .wcard-status-badge {{
    display: inline-flex; align-items: center; gap: 5px;
    padding: 4px 10px;
    border-radius: 100px;
    font-size: 11px;
    font-weight: 500;
  }}
  .wcard-status-badge.good {{ background: var(--sage-bg); color: var(--olive); }}
  .wcard-status-badge.warn {{ background: var(--gold-bg); color: var(--mocha); }}
  .wcard-status-badge.bad  {{ background: var(--terracotta-bg); color: var(--terracotta); }}
  .wcard-status-badge .dot {{ width: 5px; height: 5px; border-radius: 50%; background: currentColor; }}

  .wcard-title {{ font-size: 16px; font-weight: 600; line-height: 1.3; margin-bottom: 6px; letter-spacing: -0.01em; }}
  .wcard-subtitle {{ font-size: 12px; color: var(--soft); margin-bottom: 16px; }}
  .wcard-tags {{ display: flex; gap: 4px; flex-wrap: wrap; margin-bottom: 16px; }}
  .tag {{ display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; }}
  .tag.mocha   {{ background: var(--mocha-bg);   color: var(--mocha); }}
  .tag.sage    {{ background: var(--sage-bg);    color: var(--olive); }}
  .tag.terra   {{ background: var(--terracotta-bg); color: var(--terracotta); }}
  .tag.plum    {{ background: var(--plum-bg);    color: var(--plum); }}
  .tag.slate   {{ background: var(--slate-bg);   color: var(--slate); }}
  .tag.gold    {{ background: var(--gold-bg);    color: var(--mocha); }}

  .wcard-progress {{ margin-top: auto; padding-top: 16px; border-top: 1px dashed var(--line); }}
  .wcard-pp {{ display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px; }}
  .wcard-pp-label {{ font-size: 12px; color: var(--soft); }}
  .wcard-pp-value {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-size: 32px;
    font-weight: 200;
    letter-spacing: -0.03em;
    line-height: 1;
  }}
  .wcard-pp-value .unit {{ font-size: 14px; color: var(--soft); }}
  .wcard-bar {{ height: 5px; background: var(--bg-2); border-radius: 100px; overflow: hidden; }}
  .wcard-bar > div {{ height: 100%; border-radius: 100px; transition: width 0.8s; }}
  .wcard-bar > div.good {{ background: linear-gradient(90deg, var(--sage), #6b8e58); }}
  .wcard-bar > div.warn {{ background: linear-gradient(90deg, var(--gold-bg), var(--gold)); }}
  .wcard-bar > div.bad  {{ background: linear-gradient(90deg, var(--terracotta-bg), var(--terracotta)); }}

  .panels-grid {{ display: grid; grid-template-columns: 1.6fr 1fr; gap: 16px; margin-bottom: 40px; }}
  .panel {{
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 28px;
    box-shadow: var(--shadow-soft);
  }}
  .panel-head {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px; }}
  .panel-title {{ font-size: 17px; font-weight: 700; letter-spacing: -0.01em; }}
  .panel-sub {{ font-size: 12px; color: var(--soft); margin-bottom: 24px; }}
  .chart-legend {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 16px; font-size: 12px; color: var(--soft); }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .legend-dot {{ width: 10px; height: 10px; border-radius: 3px; }}
  .chart-wrap {{ position: relative; height: 320px; }}

  .activity {{ display: flex; flex-direction: column; gap: 2px; }}
  .activity-item {{
    display: grid; grid-template-columns: auto 1fr auto; gap: 14px;
    padding: 14px 12px; border-radius: 10px;
    transition: background 0.15s; cursor: pointer; align-items: center;
  }}
  .activity-item:hover {{ background: var(--bg-2); }}
  .activity-date {{
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-size: 26px; font-weight: 200;
    letter-spacing: -0.03em; line-height: 1;
    color: var(--ink); text-align: center; min-width: 44px;
  }}
  .activity-date small {{
    display: block; font-family: 'Pretendard Variable', sans-serif;
    font-size: 10px; color: var(--softer);
    text-transform: uppercase; letter-spacing: 0.1em;
    margin-top: 4px; font-style: normal; font-weight: 500;
  }}
  .activity-content {{ font-size: 13px; line-height: 1.5; }}
  .activity-content .title {{ font-weight: 600; color: var(--ink); margin-bottom: 2px; display: flex; gap: 6px; align-items: center; }}
  .activity-content .meta {{ color: var(--soft); font-size: 12px; }}
  .activity-qty {{
    text-align: right;
    font-family: 'Pretendard Variable', 'Pretendard', sans-serif;
    font-size: 30px; font-weight: 200;
    letter-spacing: -0.03em; line-height: 1;
    color: var(--terracotta);
  }}
  .activity-qty small {{
    display: block; font-family: 'Pretendard Variable', sans-serif;
    font-size: 10px; color: var(--softer);
    margin-top: 2px; text-transform: uppercase; letter-spacing: 0.1em;
  }}

  footer {{
    margin-top: 60px; padding: 24px 0;
    text-align: center; color: var(--softer); font-size: 12px;
    border-top: 1px dashed var(--line);
    display: flex; justify-content: space-between; align-items: center;
  }}

  @media (max-width: 1200px) {{
    .hero-stats {{ grid-template-columns: 1fr 1fr 1fr; }}
  }}
  @media (max-width: 800px) {{
    .hero-stats {{ grid-template-columns: 1fr 1fr; }}
    .panels-grid {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>
<div class="page">

<header>
  <div>
    <div class="breadcrumb">
      <a href="#">🏗️ HCNS Service</a>
      <span class="sep">/</span>
      <a href="#">CBS Improvement</a>
      <span class="sep">/</span>
      <span class="current">{site_code_display}</span>
    </div>
    <h1><span class="label">전체 진행률</span> <span class="accent">{overall_pct_str}</span></h1>
    <div class="meta-pills">
      <span class="meta-pill live"><span class="dot"></span>실시간 동기화</span>
      <span class="meta-pill"><span class="dot"></span>마지막 업데이트 {update_time}</span>
      <span class="meta-pill">📋 작업 항목 {item_count}개</span>
    </div>
  </div>
  <button class="refresh-btn" onclick="location.reload()">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
    새로고침
  </button>
</header>

{tabs_html}

{kpi_html}

<div class="section-head">
  <h2 class="section-title">
    <span class="num">i</span>
    공종별 진행 현황
  </h2>
  <div class="section-action">↻ 전체 보기</div>
</div>

{work_section_html}

<div class="section-head">
  <h2 class="section-title">
    <span class="num">ii</span>
    추이와 활동
  </h2>
</div>

<div class="panels-grid">
  <div class="panel">
    <div class="panel-head">
      <div>
        <div class="panel-title">📈 일별 작업량</div>
        <div class="panel-sub">전체 기간 일별 설치량</div>
      </div>
    </div>
    <div class="chart-legend">
      <span class="legend-item"><span class="legend-dot" style="background:var(--terracotta);"></span>일별 작업량</span>
      <span class="legend-item"><span class="legend-dot" style="background:var(--mocha-light); opacity:0.4;"></span>이동 평균 (5일)</span>
    </div>
    <div class="chart-wrap"><canvas id="trend"></canvas></div>
  </div>

  <div class="panel">
    <div class="panel-head">
      <div>
        <div class="panel-title">📅 최근 작업</div>
        <div class="panel-sub">트래킹 시트 자동 기록</div>
      </div>
    </div>
    {activity_html}
  </div>
</div>

<footer>
  <div>HCNS Service · CBS Improvement Project</div>
  <div class="mono">자동 갱신됨 · {update_time_full}</div>
</footer>

</div>

<script>
Chart.defaults.font.family = "'Pretendard Variable', 'Pretendard', sans-serif";
Chart.defaults.color = '#6b665d';

const data = {chart_data_json};
const labels = {chart_labels_json};

// 5-day moving average
const movingAvg = data.map((_, i) => {{
  const start = Math.max(0, i - 4);
  const slice = data.slice(start, i + 1);
  return slice.reduce((a, b) => a + b, 0) / slice.length;
}});

const ctx = document.getElementById('trend').getContext('2d');
const grad = ctx.createLinearGradient(0, 0, 0, 320);
grad.addColorStop(0, 'rgba(194,65,12,0.25)');
grad.addColorStop(1, 'rgba(194,65,12,0.0)');

new Chart(ctx, {{
  type: 'line',
  data: {{
    labels,
    datasets: [
      {{
        label: '이동 평균 (5일)',
        data: movingAvg,
        borderColor: '#b08968',
        borderDash: [5, 5],
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0.4,
        fill: false,
      }},
      {{
        label: '일별 작업량',
        data,
        borderColor: '#c2410c',
        backgroundColor: grad,
        fill: true,
        tension: 0.4,
        borderWidth: 2.5,
        pointRadius: 4,
        pointBackgroundColor: '#fff',
        pointBorderColor: '#c2410c',
        pointBorderWidth: 2,
        pointHoverRadius: 7,
        pointHoverBorderWidth: 3,
      }}
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        backgroundColor: '#1f1d1a',
        titleColor: '#fff',
        bodyColor: '#e6dfd1',
        padding: 12,
        cornerRadius: 8,
        titleFont: {{ weight: '600', size: 12 }},
        bodyFont: {{ size: 12 }},
        displayColors: false,
        callbacks: {{
          label: (ctx) => `${{ctx.dataset.label}}: ${{ctx.parsed.y.toFixed(1)}}`
        }}
      }}
    }},
    scales: {{
      y: {{
        grid: {{ color: '#ececec', drawTicks: false }},
        ticks: {{ padding: 8, font: {{ size: 11 }} }},
        beginAtZero: true,
        border: {{ display: false }},
      }},
      x: {{
        grid: {{ display: false }},
        ticks: {{ font: {{ size: 11 }} }},
        border: {{ color: '#d8d8d8' }},
      }}
    }},
    interaction: {{ mode: 'index', intersect: false }},
  }}
}});
</script>

</body>
</html>
"""


def render_dashboard(data: DashboardData, active_site_code: Optional[str] = None) -> str:
    """
    대시보드 HTML 전체 생성.
    
    Args:
        data: extract_dashboard_data 결과.
        active_site_code: 메인으로 표시할 시트 코드. None이면 진행률 가장 높은 곳.
    
    Returns:
        완성된 HTML 문자열.
    """
    # 활성 시트 결정 (기본: 누적 작업이 가장 많은 시트)
    if active_site_code is None:
        non_empty = [s for s in data.sites if s.total_actual > 0]
        if non_empty:
            active_site = max(non_empty, key=lambda s: s.total_actual)
        else:
            active_site = data.sites[0] if data.sites else None
    else:
        active_site = next(
            (s for s in data.sites if s.site_code == active_site_code),
            data.sites[0] if data.sites else None,
        )
    
    if active_site is None:
        return "<html><body><h1>데이터 없음</h1></body></html>"
    
    # 부분 HTML 렌더링
    tabs_html = render_tabs(data.sites, active_site.site_code)
    kpi_html = render_kpi_section(active_site, data)
    work_section_html = render_work_section(active_site)
    # 현장별로 분리된 활동/추이 사용 (없으면 전체 합산으로 폴백)
    site_activity = active_site.recent_activity or data.recent_activity
    site_trend = active_site.daily_trend or data.daily_trend
    activity_html = render_activity_section(site_activity)
    chart_labels_json, chart_data_json = render_chart_data(site_trend)
    
    # 메타 정보
    overall_pct_str = f"{active_site.overall_pct * 100:.1f}%"
    update_time = data.generated_at.strftime("%H:%M")
    update_time_full = data.generated_at.strftime("%Y-%m-%d %H:%M")
    site_code_display = "기타" if active_site.site_code == "None PJT" else active_site.site_code
    
    return HTML_TEMPLATE.format(
        site_code=active_site.site_code,
        site_code_display=site_code_display,
        overall_pct_str=overall_pct_str,
        update_time=update_time,
        update_time_full=update_time_full,
        item_count=len(active_site.items),
        tabs_html=tabs_html,
        kpi_html=kpi_html,
        work_section_html=work_section_html,
        activity_html=activity_html,
        chart_data_json=chart_data_json,
        chart_labels_json=chart_labels_json,
    )


if __name__ == "__main__":
    import sys
    from extract_data import extract_dashboard_data
    
    pass
