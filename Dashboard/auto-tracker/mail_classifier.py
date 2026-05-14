"""
Claude API(Sonnet 4.6) 기반 업무 메일 분류기.

입력: 메일 1통의 메타데이터(발신자/제목/수신일시/첨부)와 본문 텍스트.
출력: 표준화된 분류 결과 dict (Mail_Log.xlsx에 그대로 들어갈 값들).

특징:
- anthropic SDK 사용. 키는 .env의 ANTHROPIC_API_KEY 또는 환경변수에서 읽음.
- 프롬프트 캐시: system 프롬프트(스키마 설명)를 cache_control로 캐시.
- 오류 시 신뢰도='낮음' + 상태='검토필요'로 fallback (스펙 준수).
"""
from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass
from typing import Optional

from mail_log_schema import (
    WORK_TYPES, PRIORITIES, STATUSES, CONFIDENCES, MAIL_LOG_COLUMNS,
)

logger = logging.getLogger("mail_classifier")


CLASSIFIER_MODEL = "claude-sonnet-4-6"

# 본문이 너무 길면 토큰 폭증하므로 잘라냄 (필요시 조정)
MAX_BODY_CHARS = 8000


# ============================================================
# 광고/비업무 1차 필터 (API 호출 전 — 비용 절감)
# ============================================================
_AD_FROM_PATTERNS = [
    r"no[-_.]?reply", r"do[-_.]?not[-_.]?reply", r"newsletter@",
    r"marketing@", r"info@.*ad", r"notice@", r"alert@", r"alarm@",
    r"notification@", r"news@", r"mailer@",
]
_AD_SUBJECT_PATTERNS = [
    r"\[광고\]", r"\(광고\)", r"AD\]", r"\[AD\]",
    r"뉴스레터", r"newsletter", r"unsubscribe", r"수신거부",
    r"이벤트 안내", r"할인 쿠폰", r"무료체험",
]


def is_obvious_non_business(sender_email: str, subject: str) -> tuple[bool, str]:
    """발신자/제목만으로 명확히 광고/알림으로 판단되면 (True, 사유) 반환."""
    email_l = (sender_email or "").lower()
    for pat in _AD_FROM_PATTERNS:
        if re.search(pat, email_l):
            return True, f"시스템 발신자 패턴({pat})"
    subj = subject or ""
    for pat in _AD_SUBJECT_PATTERNS:
        if re.search(pat, subj, flags=re.IGNORECASE):
            return True, f"광고/뉴스레터 제목 패턴({pat})"
    return False, ""


# ============================================================
# 분류 결과 정규화 유틸
# ============================================================
def _safe_pick(value, allowed: list[str], default: str) -> str:
    if value is None:
        return default
    s = str(value).strip()
    return s if s in allowed else default


def _coerce_number(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s:
        return ""
    # 콤마, 통화 기호 제거
    s2 = re.sub(r"[,\s]", "", s)
    s2 = re.sub(r"[₩원$￦KRWUSD]", "", s2, flags=re.IGNORECASE)
    try:
        if "." in s2:
            return float(s2)
        return int(s2)
    except ValueError:
        return ""


def _coerce_date(v) -> str:
    if not v:
        return ""
    s = str(v).strip()
    # YYYY-MM-DD 검증
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return ""


# ============================================================
# 시스템 프롬프트 (캐시됨)
# ============================================================
SYSTEM_PROMPT = """\
당신은 한국 회사 직원의 받은편지함에서 업무 메일을 분류·요약하는 어시스턴트입니다.

다음 JSON 스키마에 정확히 맞춰 1개의 JSON 객체만 출력하세요. 마크다운 코드펜스나 설명 문장은 절대 출력하지 마세요.

스키마:
{
  "is_business": true | false,
  "non_business_reason": "광고|뉴스레터|자동알림|스팸|단순참조|" (is_business=false일 때만)
  "회사명": string,                  // 발신자가 속한 회사/고객사 추정. 모르면 ""
  "업무유형": "견적"|"주문"|"계약"|"정산"|"CS"|"일정"|"보고"|"기타",
  "핵심내용요약": string,            // 1~3문장으로 요약
  "요청사항": string,                // 발신자가 요구하는 액션. 없으면 ""
  "금액": number | "",               // 한국원화 기준 숫자만(콤마/₩/원 빼고). 없으면 ""
  "수량": number | "",               // 단위 없는 정수. 없으면 ""
  "납기일": "YYYY-MM-DD" | "",       // 마감/납기 날짜. 없으면 ""
  "담당자": string,                  // 메일 본문/서명에서 추정. 없으면 ""
  "우선순위": "긴급"|"높음"|"보통"|"낮음",
  "신뢰도": "높음"|"중간"|"낮음",   // 추출 결과 자체에 대한 본인 확신도
  "검토필요": true | false,          // 사람 확인이 필요하다고 판단되면 true
  "비고": string                     // 추가 메모(추정값 표시 등). 없으면 ""
}

규칙:
- 업무 요청/견적/주문/계약/정산/CS/일정/보고/첨부파일 전달 등 실제 업무 데이터가 포함된 메일이면 is_business=true.
- 광고, 뉴스레터, 자동 알림, 스팸, 단순 참조(FYI/CC만)면 is_business=false.
- 애매하면 is_business=true + 검토필요=true + 신뢰도="낮음" 으로 둡니다. 절대 임의 제외 금지.
- 우선순위 단서: "긴급", "ASAP", "오늘 내", "급" → 긴급. "내일까지", "이번 주" → 높음. 명시 없으면 "보통".
- 본문/제목에 명시 없는 값은 추측해서 넣지 말고 ""로 두세요. 추정한 값이 있으면 비고에 "추정: <필드>" 라고 명시.
- 금액·수량·날짜는 숫자/표준형식으로만. "약 1억" 같은 모호한 표현은 "" + 비고에 원문 인용.
"""


@dataclass
class ClassifierResult:
    is_business: bool
    fields: dict  # MAIL_LOG_COLUMNS 중 분류기가 채울 수 있는 키들의 값
    non_business_reason: str = ""
    raw_response: str = ""


# ============================================================
# 메인 분류 함수
# ============================================================
def classify_mail(
    *,
    api_key: str,
    sender_name: str,
    sender_email: str,
    subject: str,
    received_at: str,
    body_text: str,
    attachments: list[str],
) -> ClassifierResult:
    """Claude API를 호출해 메일을 분류·추출한다.

    실패 시 신뢰도='낮음' + 상태='검토필요' fallback 결과 반환.
    """
    # 1차 광고 필터
    is_ad, ad_reason = is_obvious_non_business(sender_email, subject)
    if is_ad:
        return ClassifierResult(
            is_business=False,
            non_business_reason=ad_reason,
            fields={},
        )

    try:
        from anthropic import Anthropic
    except ImportError:
        logger.error("anthropic 패키지가 설치되어 있지 않습니다. pip install anthropic")
        return _fallback_result("anthropic 패키지 미설치")

    if not api_key:
        return _fallback_result("ANTHROPIC_API_KEY 미설정")

    client = Anthropic(api_key=api_key)

    # 본문 길이 제한
    body_excerpt = body_text or ""
    if len(body_excerpt) > MAX_BODY_CHARS:
        body_excerpt = body_excerpt[:MAX_BODY_CHARS] + "\n...(본문 일부 생략됨)..."

    att_str = ", ".join(attachments) if attachments else "(없음)"

    user_msg = (
        f"수신일시: {received_at}\n"
        f"발신자: {sender_name} <{sender_email}>\n"
        f"제목: {subject}\n"
        f"첨부파일: {att_str}\n"
        f"--- 본문 시작 ---\n{body_excerpt}\n--- 본문 끝 ---"
    )

    try:
        resp = client.messages.create(
            model=CLASSIFIER_MODEL,
            max_tokens=1024,
            system=[
                {
                    "type": "text",
                    "text": SYSTEM_PROMPT,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[{"role": "user", "content": user_msg}],
        )
    except Exception as exc:
        logger.error("Claude API 호출 실패: %s", exc)
        return _fallback_result(f"API 호출 실패: {exc}")

    # 응답 텍스트 추출
    raw = ""
    for block in resp.content:
        if getattr(block, "type", None) == "text":
            raw += block.text

    # JSON 파싱 (코드펜스가 섞여 들어와도 추출)
    parsed = _extract_json(raw)
    if parsed is None:
        logger.warning("JSON 파싱 실패. raw=%s", raw[:200])
        return _fallback_result("응답 JSON 파싱 실패", raw_response=raw)

    return _normalize_result(parsed, raw)


# ============================================================
# JSON 추출/정규화
# ============================================================
def _extract_json(text: str) -> Optional[dict]:
    if not text:
        return None
    # 코드펜스 제거
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = m.group(1) if m else text
    # 첫 { 부터 마지막 } 까지
    start = candidate.find("{")
    end = candidate.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        return json.loads(candidate[start:end + 1])
    except json.JSONDecodeError:
        return None


def _normalize_result(parsed: dict, raw: str) -> ClassifierResult:
    is_biz = bool(parsed.get("is_business", True))
    if not is_biz:
        return ClassifierResult(
            is_business=False,
            non_business_reason=str(parsed.get("non_business_reason", "")),
            fields={},
            raw_response=raw,
        )

    review_needed = bool(parsed.get("검토필요", False))
    confidence = _safe_pick(parsed.get("신뢰도"), CONFIDENCES, "중간")
    if confidence == "낮음":
        review_needed = True

    status = "검토필요" if review_needed else "신규"

    fields = {
        "회사명": str(parsed.get("회사명", "") or ""),
        "업무유형": _safe_pick(parsed.get("업무유형"), WORK_TYPES, "기타"),
        "핵심내용요약": str(parsed.get("핵심내용요약", "") or ""),
        "요청사항": str(parsed.get("요청사항", "") or ""),
        "금액": _coerce_number(parsed.get("금액")),
        "수량": _coerce_number(parsed.get("수량")),
        "납기일": _coerce_date(parsed.get("납기일")),
        "담당자": str(parsed.get("담당자", "") or ""),
        "우선순위": _safe_pick(parsed.get("우선순위"), PRIORITIES, "보통"),
        "상태": status,
        "신뢰도": confidence,
        "비고": str(parsed.get("비고", "") or ""),
    }
    return ClassifierResult(
        is_business=True,
        fields=fields,
        raw_response=raw,
    )


def _fallback_result(reason: str, raw_response: str = "") -> ClassifierResult:
    """API 호출 실패 시 — 메일을 버리지 않고 검토필요로 기록."""
    return ClassifierResult(
        is_business=True,
        fields={
            "회사명": "",
            "업무유형": "기타",
            "핵심내용요약": "",
            "요청사항": "",
            "금액": "",
            "수량": "",
            "납기일": "",
            "담당자": "",
            "우선순위": "보통",
            "상태": "검토필요",
            "신뢰도": "낮음",
            "비고": f"자동분류 실패: {reason}",
        },
        raw_response=raw_response,
    )
