# -*- coding: utf-8 -*-
"""환불·취소 감지 공통 장치 (전 플랫폼).

환불·취소 정책(2026-07-29 결정, docs/PLATFORMS.md)
- partially_refunded(부분환불·클레임 보상) → 전액 반영.
  물품 반환 없는 클레임 변상금은 과세표준에서 공제하지 않음 (부가46015-2537, 1996.11.28.)
- 전액환불·반품·취소·미배송 → 미반영 + 사유별 집계

이 모듈은 위 정책을 아는 파서(쇼피파이·라자다 주문내역)가 없는 입력에서도
환불성 자료가 조용히 매출에 섞이거나 빠지지 않도록 감시합니다.

1) looks_refundish  — 상태 문자열의 환불성 키워드 판정 (파서 공용 키워드)
2) negative_warning — 음수 금액 행 요약 경고문 생성
3) scan_pdf_text    — 소포수령증 PDF 원문에서 환불성 표기 스캔

감지 결과는 금액을 바꾸지 않고 결과 dict의 refund_warnings 로 보고만 합니다.
예외는 음수 금액 행 하나 — 이는 매출 합계에서 제외하고 미반영/검토 행으로 돌립니다
(음수 행을 그대로 더하면 환불이 소리 없이 차감돼 정책과 어긋나기 때문입니다).
"""

from __future__ import annotations

# 상태 문자열 판정용 — 쇼피파이 financial_status, 라자다 status 등에 사용
REFUND_STATUS_KEYWORDS = (
    "refund", "return", "cancel", "void", "chargeback",
    "환불", "반품", "취소",
)

# PDF 원문 스캔용 — 소포수령증 양식은 배송분만 기재되는 것이 전제라서,
# 이런 표기가 보이면 양식이 바뀌었거나 환불 내역이 섞였다는 신호입니다.
# ('return'/'cancel' 영단어는 주소·안내문 오탐 여지가 있어 제외합니다.)
PDF_SCAN_KEYWORDS = ("환불", "반품", "refund", "chargeback")


def looks_refundish(text) -> bool:
    """상태 문자열에 환불·반품·취소성 키워드가 있는지 판정합니다."""
    t = str(text or "").strip().lower()
    if not t:
        return False
    return any(k in t for k in REFUND_STATUS_KEYWORDS)


def negative_warning(platform: str, count: int, total_by_currency: dict) -> str:
    """음수 금액 행 요약 경고문. total_by_currency: {'USD': -22.75, ...}"""
    amounts = ", ".join(
        f"{v:,.2f} {cur}" for cur, v in sorted(total_by_currency.items())
    )
    return (
        f"{platform}: 음수 금액 {count}건 발견 ({amounts}) — 환불·조정으로 보입니다. "
        "매출 합계에서 제외하고 미반영(검토) 행으로 표시했습니다. 원본을 확인해 주세요."
    )


def scan_pdf_text(platform: str, text: str) -> list:
    """소포수령증 PDF 원문에서 환불성 표기를 찾아 경고 목록을 반환합니다.

    이 파서들은 환불 처리를 지원하지 않으므로(배송분만 기재되는 양식),
    표기가 발견되면 숫자는 그대로 두고 검토하라는 경고만 냅니다.
    """
    lower = str(text or "").lower()
    found = [k for k in PDF_SCAN_KEYWORDS if k in lower]
    if not found:
        return []
    return [
        f"{platform}: 수령증 PDF에서 환불성 표기({', '.join(found)})가 발견됐습니다. "
        "이 양식은 배송분만 집계하므로 환불 내역이 섞였는지 원본 PDF를 확인해 주세요."
    ]
