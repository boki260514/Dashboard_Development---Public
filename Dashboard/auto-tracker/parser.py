"""
카톡 메시지 파싱 모듈.

카톡으로 받는 작업 일보 텍스트를 분석하여 구조화된 데이터로 변환합니다.

지원하는 메시지 형식 예시:
    [ 쿠팡 창원 1센터 ]
    작업일자 : 04.30
    작업시간 : 04:30 - 06:30
    작업인원 : 총원 13명
                      관리자 1, 작업자 12
    작업내용 :
    - PB #1 Line
       => Gap Cover
             -> NTP : 318~366 《 49ea 》
    - PB #2 Line
       => Gap Cover
             -> TP : 192~235 《 44ea 》
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

from config import SITE_MAPPING, WORK_TYPE_KEYWORDS


@dataclass
class RawDetail:
    """파싱된 개별 라인의 원본 정보 (TP/NTP 매칭에 사용)."""
    tp_or_ntp: str         # "TP" / "NTP" / "" (없으면)
    number_range: str      # 예: "253~288"
    quantity: float        # 예: 36


@dataclass
class WorkItem:
    """단일 작업 항목 (예: PB#1 - Gap Cover - 49ea)."""
    line_raw: str          # 원본 라인명 (예: "PB #1 Line")
    work_type_raw: str     # 원본 작업명 (예: "Gap Cover")
    line_normalized: str   # 정규화된 라인 키 (예: "pb#1line")
    work_type: str         # 작업 종류 키 (예: "gapcover")
    quantity: float        # 작업량 (TP/NTP 세트 매칭 후 최종값)
    detail: str = ""       # 상세 설명 (예: "TP 192~235")
    raw_details: list[RawDetail] = field(default_factory=list)
    # 후처리 시 worker_count 분배 결과를 저장
    workers: int = 0


@dataclass
class DailyReport:
    """하루 작업 일보."""
    site_raw: str          # 원본 현장명 (예: "쿠팡 창원 1센터")
    site_sheet: str        # 매칭된 시트명 (예: "CHW1FC")
    work_date: Optional[date] = None  # 작업 일자
    start_time: Optional[str] = None  # 작업 시작 시간
    end_time: Optional[str] = None    # 작업 종료 시간
    work_hours: Optional[float] = None  # 작업 시간 (자동 계산)
    total_workers: Optional[int] = None  # 총 인원
    manager_count: Optional[int] = None  # 관리자 수
    worker_count: Optional[int] = None   # 작업자 수
    items: list[WorkItem] = field(default_factory=list)  # 작업 항목들
    raw_message: str = ""  # 원본 메시지


def normalize(text: str) -> str:
    """
    문자열을 매칭용 정규 형태로 변환합니다.

    - 모든 공백 제거
    - 소문자로 변환
    - 특수 따옴표/괄호 제거

    Args:
        text: 원본 문자열.

    Returns:
        정규화된 문자열.
    """
    if not text:
        return ""
    # 공백, 탭, 줄바꿈 모두 제거
    s = re.sub(r"\s+", "", text)
    # 특수 괄호류 제거
    s = re.sub(r"[《》〈〉「」『』【】\[\](){}]", "", s)
    return s.lower()


def find_site(message: str) -> tuple[str, str] | tuple[None, None]:
    """
    메시지에서 현장명을 찾아 시트명으로 매핑합니다.

    Args:
        message: 카톡 메시지 전문.

    Returns:
        (원본 현장명, 시트명) 튜플. 못 찾으면 (None, None).
    """
    # [ ... ] 안의 내용을 우선 검사
    bracket_match = re.search(r"\[\s*([^\]]+?)\s*\]", message)
    candidates = []
    if bracket_match:
        candidates.append(bracket_match.group(1))

    # 추가로 메시지의 첫 3줄도 후보에 포함
    first_lines = message.strip().split("\n")[:3]
    candidates.extend(first_lines)

    for candidate in candidates:
        normalized = normalize(candidate)
        for key, sheet in SITE_MAPPING.items():
            if key in normalized:
                return candidate.strip(), sheet

    return None, None


def parse_date(message: str) -> Optional[date]:
    """
    메시지에서 작업 일자를 추출합니다.

    지원 형식:
        - "작업일자 : 04.30"
        - "작업일자: 4/30"
        - "작업일자 04월 30일"
        - "2026.04.30"
        - "2026-04-30"

    Args:
        message: 카톡 메시지 전문.

    Returns:
        date 객체. 못 찾으면 None.
    """
    today = date.today()

    # 1. 전체 날짜 (YYYY.MM.DD 또는 YYYY-MM-DD)
    m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", message)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    # 2. 작업일자 라벨이 있는 경우 (MM.DD / MM/DD / MM월 DD일)
    m = re.search(
        r"작업\s*일자\s*[:\-]?\s*"
        r"(\d{1,2})\s*[.\-/월]\s*(\d{1,2})",
        message,
    )
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            return date(today.year, month, day)
        except ValueError:
            pass

    # 3. 라벨 없이 MM.DD 형식
    m = re.search(r"(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)", message)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            return date(today.year, month, day)
        except ValueError:
            pass

    return None


def parse_work_time(message: str) -> tuple[Optional[str], Optional[str], Optional[float]]:
    """
    메시지에서 작업 시간을 추출합니다.

    형식: "작업시간 : 04:30 - 06:30"

    Args:
        message: 카톡 메시지 전문.

    Returns:
        (시작시간, 종료시간, 총시간(시)) 튜플.
    """
    m = re.search(
        r"작업\s*시간\s*[:\-]?\s*"
        r"(\d{1,2}):(\d{2})\s*[-~]\s*(\d{1,2}):(\d{2})",
        message,
    )
    if not m:
        return None, None, None

    start_h, start_m = int(m.group(1)), int(m.group(2))
    end_h, end_m = int(m.group(3)), int(m.group(4))
    start_time = f"{start_h:02d}:{start_m:02d}"
    end_time = f"{end_h:02d}:{end_m:02d}"

    # 시간 차 계산 (분 단위로 → 시간으로)
    start_total = start_h * 60 + start_m
    end_total = end_h * 60 + end_m
    if end_total < start_total:
        # 자정을 넘긴 경우
        end_total += 24 * 60
    diff_hours = round((end_total - start_total) / 60, 2)

    return start_time, end_time, diff_hours


def parse_workers(message: str) -> tuple[Optional[int], Optional[int], Optional[int]]:
    """
    메시지에서 작업 인원을 추출합니다.

    형식:
        작업인원 : 총원 13명
                          관리자 1, 작업자 12

    Args:
        message: 카톡 메시지 전문.

    Returns:
        (총원, 관리자수, 작업자수) 튜플.
    """
    total = manager = worker = None

    # 총원
    m = re.search(r"총\s*원\s*[:\-]?\s*(\d+)\s*명", message)
    if m:
        total = int(m.group(1))

    # 관리자
    m = re.search(r"관\s*리\s*자\s*[:\-]?\s*(\d+)", message)
    if m:
        manager = int(m.group(1))

    # 작업자
    m = re.search(r"작\s*업\s*자\s*[:\-]?\s*(\d+)", message)
    if m:
        worker = int(m.group(1))

    # 총원이 없으면 관리자+작업자 합산
    if total is None and manager is not None and worker is not None:
        total = manager + worker

    return total, manager, worker


def parse_work_items(message: str) -> list[WorkItem]:
    """
    메시지의 작업내용 섹션에서 개별 작업 항목들을 추출합니다.

    형식:
        - PB #1 Line
           => Gap Cover
                 -> NTP : 318~366 《 49ea 》
                 -> TP : 192~235 《 44ea 》

    각 수량 라인에서 TP/NTP 구분, 번호 범위, 수량을 모두 추출하여
    raw_details에 저장합니다. 이후 post_process_items에서 TP/NTP
    매칭으로 세트 단위 수량을 계산합니다.

    Args:
        message: 카톡 메시지 전문.

    Returns:
        WorkItem 리스트 (각각에 raw_details가 채워져 있음).
    """
    items: list[WorkItem] = []

    current_line = None
    current_work_type = None

    for raw_line in message.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        # 라인 식별: "- PB #1 Line"
        line_match = re.match(
            r"^[-•]\s*([A-Z]+\s*#?\s*\d*\s*[Ll]ine?)",
            line,
        )
        if line_match:
            current_line = line_match.group(1).strip()
            current_work_type = None
            continue

        # 작업 종류: "=> Gap Cover" 또는 "=> 하네스 케이블"
        work_match = re.match(r"^=+>?\s*(.+)$", line)
        if work_match and current_line:
            work_text = work_match.group(1).strip()
            for wt_key in ["gapcover", "gapplate", "cable", "cover"]:
                keywords = WORK_TYPE_KEYWORDS.get(wt_key, [])
                for kw in keywords:
                    if kw.lower() in work_text.lower():
                        current_work_type = wt_key
                        break
                if current_work_type:
                    break
            continue

        # 수량 라인 패턴 분석
        if not (current_line and current_work_type):
            continue

        # 패턴: TP/NTP 라벨 + 번호 범위 + 수량
        # 예: "-> NTP : 318~366 《 49ea 》"
        # 예: "    253~288 《 36ea 》" (앞 줄에서 NTP가 이미 나온 경우)
        labeled_match = re.search(
            r"(NTP|TP|NT)\s*[:：]\s*"
            r"(\d+\s*[~\-]\s*\d+)"
            r".*?[《(\[]\s*(\d+(?:\.\d+)?)\s*ea",
            line,
            re.IGNORECASE,
        )
        if labeled_match:
            tp_label = labeled_match.group(1).upper()
            num_range = labeled_match.group(2).replace(" ", "")
            qty = float(labeled_match.group(3))
            _add_or_create_item(
                items, current_line, current_work_type,
                tp_label, num_range, qty,
            )
            continue

        # 라벨이 없는 추가 라인 (예: "289~317 《 29ea 》")
        # 직전 항목의 TP/NTP 라벨을 그대로 사용
        unlabeled_match = re.search(
            r"(\d+\s*[~\-]\s*\d+)"
            r".*?[《(\[]\s*(\d+(?:\.\d+)?)\s*ea",
            line,
            re.IGNORECASE,
        )
        if unlabeled_match:
            num_range = unlabeled_match.group(1).replace(" ", "")
            qty = float(unlabeled_match.group(2))
            # 가장 최근에 추가된 같은 라인+작업의 마지막 TP/NTP 라벨 가져오기
            last_label = _find_last_label(
                items, current_line, current_work_type,
            )
            if last_label is None:
                # 라벨을 모르면 그냥 빈 값으로 추가 (짝 없는 항목으로 처리됨)
                last_label = ""
            _add_or_create_item(
                items, current_line, current_work_type,
                last_label, num_range, qty,
            )
            continue

        # 패턴 2: 라벨 없이 그냥 (1ea) 같은 형식 (하네스 케이블 등)
        plain_match = re.search(
            r"[《(\[]\s*(\d+(?:\.\d+)?)\s*ea\s*[》)\]]",
            line,
            re.IGNORECASE,
        )
        if plain_match:
            qty = float(plain_match.group(1))
            _add_or_create_item(
                items, current_line, current_work_type,
                "", "", qty,
            )

    return items


def _find_last_label(
    items: list[WorkItem],
    line: str,
    work_type: str,
) -> Optional[str]:
    """이전에 등록된 같은 라인+작업의 마지막 TP/NTP 라벨을 찾습니다."""
    line_norm = normalize(line)
    for item in reversed(items):
        if (item.line_normalized == line_norm and
                item.work_type == work_type):
            for d in reversed(item.raw_details):
                if d.tp_or_ntp:
                    return d.tp_or_ntp
    return None


def _add_or_create_item(
    items: list[WorkItem],
    line: str,
    work_type: str,
    tp_label: str,
    num_range: str,
    qty: float,
) -> None:
    """기존 WorkItem이 있으면 raw_details에 추가, 없으면 새로 생성."""
    line_norm = normalize(line)
    for item in items:
        if (item.line_normalized == line_norm and
                item.work_type == work_type):
            item.raw_details.append(RawDetail(
                tp_or_ntp=tp_label,
                number_range=num_range,
                quantity=qty,
            ))
            return
    # 새 WorkItem
    items.append(WorkItem(
        line_raw=line,
        work_type_raw=work_type,
        line_normalized=line_norm,
        work_type=work_type,
        quantity=0.0,  # post_process에서 계산됨
        raw_details=[RawDetail(
            tp_or_ntp=tp_label,
            number_range=num_range,
            quantity=qty,
        )],
    ))


def post_process_items(items: list[WorkItem]) -> list[WorkItem]:
    """
    각 WorkItem의 raw_details를 TP/NTP 세트 매칭으로 합산합니다.

    규칙:
      - 같은 번호 범위에 TP와 NTP 둘 다 있으면 → 한 세트 (그 수량 그대로)
      - 한쪽만 있으면 (TP만 또는 NTP만) → 수량을 절반으로 (반쪽 세트)
      - 라벨이 없으면 (예: 하네스 케이블 1ea) → 그대로 합산

    예시:
      TP   253~288 36ea
      NTP  253~288 36ea  ← 같은 번호 → 36 (1세트)
      NTP  289~317 29ea  ← 짝 없음   → 14.5 (반쪽 세트)
      = 합계 50.5

    Args:
        items: parse_work_items에서 반환된 raw_details가 채워진 항목들.

    Returns:
        quantity가 세트 단위로 계산된 항목 리스트.
    """
    for item in items:
        item.quantity = _calculate_set_quantity(item.raw_details)
        item.detail = _build_detail_text(item.raw_details)
    return items


def _calculate_set_quantity(details: list[RawDetail]) -> float:
    """
    raw_details를 TP/NTP 매칭으로 세트 수량을 계산합니다.

    Args:
        details: 한 작업의 모든 raw 정보들.

    Returns:
        세트 단위로 계산된 총 수량.
    """
    # 라벨 없는 항목들은 그대로 합산 (예: 하네스 케이블 (1ea))
    # 라벨 있는 항목들만 번호 범위로 그룹핑하여 매칭
    by_range: dict[str, dict[str, float]] = {}
    no_label_total = 0.0

    for d in details:
        # TP/NTP 라벨이 없으면 그대로 합산 (번호 범위 무시)
        if not d.tp_or_ntp:
            no_label_total += d.quantity
            continue
        key = d.number_range or f"_no_range_{id(d)}"
        if key not in by_range:
            by_range[key] = {}
        by_range[key][d.tp_or_ntp] = d.quantity

    total = no_label_total
    for num_range, label_map in by_range.items():
        # TP와 NTP/NT 둘 다 있는지 확인
        has_tp = "TP" in label_map
        has_ntp = "NTP" in label_map or "NT" in label_map

        # TP와 NTP 모두 있으면 세트 → 둘 중 하나의 값 (보통 같음)
        if has_tp and has_ntp:
            tp_val = label_map.get("TP", 0)
            ntp_val = label_map.get("NTP") or label_map.get("NT") or 0
            if tp_val == ntp_val:
                total += tp_val
            else:
                # 다르면 작은 값을 세트로 보고, 차이는 반쪽
                pair = min(tp_val, ntp_val)
                leftover = abs(tp_val - ntp_val)
                total += pair + leftover / 2
        else:
            # 한쪽만 있음 (TP만 또는 NTP만) → 반으로 나누기
            for v in label_map.values():
                total += v / 2

    return total


def _build_detail_text(details: list[RawDetail]) -> str:
    """raw_details를 사람이 읽기 좋은 한 줄로 변환합니다."""
    parts = []
    for d in details:
        if d.tp_or_ntp and d.number_range:
            parts.append(f"{d.tp_or_ntp} {d.number_range} {d.quantity}ea")
        elif d.number_range:
            parts.append(f"{d.number_range} {d.quantity}ea")
        else:
            parts.append(f"{d.quantity}ea")
    return " + ".join(parts)


def distribute_workers(
    items: list[WorkItem],
    total_workers: int,
) -> None:
    """
    총 인원을 작업량 비례로 각 항목에 정수로 분배합니다.

    분배 알고리즘 (Largest Remainder Method):
      1. 각 항목의 비율 계산 (quantity / total_quantity)
      2. 비율 × 총인원 = 이상적 분배값 (소수)
      3. 소수의 정수부 먼저 할당
      4. 남은 인원은 소수부가 큰 항목부터 +1씩 분배
      5. 합계가 정확히 total_workers가 되도록 보장

    Args:
        items: WorkItem 리스트 (각각의 quantity가 계산된 상태).
        total_workers: 분배할 총 인원.
    """
    if not items or total_workers is None or total_workers <= 0:
        return

    total_qty = sum(item.quantity for item in items)
    if total_qty <= 0:
        # 작업량이 모두 0이면 균등 분배
        base = total_workers // len(items)
        remainder = total_workers - base * len(items)
        for i, item in enumerate(items):
            item.workers = base + (1 if i < remainder else 0)
        return

    # 1단계: 이상적 분배값 계산
    ideal = [(item.quantity / total_qty) * total_workers for item in items]

    # 2단계: 정수부 먼저 할당
    floors = [int(v) for v in ideal]
    remainders = [(ideal[i] - floors[i], i) for i in range(len(items))]

    # 3단계: 남은 인원을 소수부 큰 순서로 +1
    remaining = total_workers - sum(floors)
    # 소수부 내림차순 정렬
    remainders.sort(key=lambda x: x[0], reverse=True)

    for k in range(remaining):
        _, idx = remainders[k % len(remainders)]
        floors[idx] += 1

    # 결과 적용
    for item, w in zip(items, floors):
        item.workers = w


def parse_message(message: str) -> DailyReport:
    """
    카톡 메시지 전체를 파싱하여 DailyReport 객체를 반환합니다.

    Args:
        message: 카톡 메시지 전문.

    Returns:
        파싱된 DailyReport.

    Raises:
        ValueError: 현장명을 찾을 수 없는 경우.
    """
    site_raw, site_sheet = find_site(message)
    if not site_sheet:
        raise ValueError(
            "메시지에서 현장명을 찾을 수 없습니다. "
            "config.py의 SITE_MAPPING을 확인하세요."
        )

    work_date = parse_date(message)
    start_time, end_time, work_hours = parse_work_time(message)
    total, manager, worker = parse_workers(message)

    items = parse_work_items(message)
    items = post_process_items(items)

    # 인원 분배 (작업량 비례, 정수)
    if total is not None:
        distribute_workers(items, total)

    return DailyReport(
        site_raw=site_raw,
        site_sheet=site_sheet,
        work_date=work_date,
        start_time=start_time,
        end_time=end_time,
        work_hours=work_hours,
        total_workers=total,
        manager_count=manager,
        worker_count=worker,
        items=items,
        raw_message=message,
    )
