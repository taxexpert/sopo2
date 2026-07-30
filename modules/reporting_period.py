# -*- coding: utf-8 -*-
"""신고기간 분류 — 업로드 자료에서 '신고 해당분'만 남기고 나머지는 사유와 함께 보존합니다.

파서 결과를 소비하는 지점이 매출집계(`excel_writer`)와 신고서류(`extra_docs`)로
갈라져 있어, 양쪽에 필터를 따로 구현하면 두 산출물의 합계가 어긋납니다.
그래서 파싱 직후 **한 번만** 이 모듈을 통과시키고, 걸러진 결과를 모든 소비자
(환율 수집 범위·파일명 라벨·매출집계·신고서류)가 공유합니다.

기준일은 선적일자(기적일)입니다 (docs/POLICY.md).

분류:
    해당분          시작일 이상, 종료일 이하 (양 끝 포함)
    이전기간        시작일 전
    미도래          종료일 후
    날짜없음        기준일을 확인할 수 없음
    기간중첩-분할불가  집계자료가 신고기간 경계를 걸침 → 생성 중단

라자다 소포수령증 PDF·이베이 월집계·큐텐처럼 건별 날짜가 없는 집계자료가 경계를
걸치면 금액을 비례 안분하지 않고 생성을 중단합니다 (docs/POLICY.md).
"""

from __future__ import annotations

import calendar
import re
from copy import deepcopy

IN_PERIOD = "해당분"
BEFORE_PERIOD = "이전기간"
AFTER_PERIOD = "미도래"
NO_DATE = "날짜없음"
STRADDLE = "기간중첩-분할불가"

EXCLUDED_CATEGORIES = (BEFORE_PERIOD, AFTER_PERIOD, NO_DATE)


class PeriodStraddleError(RuntimeError):
    """분할할 수 없는 집계자료가 신고기간 경계를 걸쳐 생성을 중단합니다."""

    def __init__(self, blocked):
        self.blocked = list(blocked)
        lines = [
            f"  · {b['platform']} {b['label']} — 자료기간 {b['span']}"
            for b in self.blocked
        ]
        super().__init__(
            "신고기간 경계를 걸치는 집계자료가 있어 생성을 중단했습니다.\n"
            + "\n".join(lines)
            + "\n건별 날짜가 없는 집계자료는 금액을 나눌 수 없습니다. "
            "신고기간에 맞는 자료를 다시 받아 올리거나, 신고기간을 자료 범위에 맞춰 주세요."
        )


# ── 날짜 유틸 ────────────────────────────────────────────────────

def date_to_int(value):
    """'2026-01-31' · '2026.01.31' · '20260131' → 20260131. 실패하면 None."""
    digits = re.sub(r"\D", "", str(value or ""))[:8]
    return int(digits) if len(digits) == 8 else None


def int_to_date(value):
    text = str(value or "")
    return f"{text[:4]}-{text[4:6]}-{text[6:8]}" if len(text) == 8 else ""


def month_span(month_value):
    """'2026-03' → (20260301, 20260331). 실패하면 (None, None)."""
    digits = re.sub(r"\D", "", str(month_value or ""))[:6]
    if len(digits) != 6:
        return None, None
    year, month = int(digits[:4]), int(digits[4:6])
    if not 1 <= month <= 12:
        return None, None
    last = calendar.monthrange(year, month)[1]
    return year * 10000 + month * 100 + 1, year * 10000 + month * 100 + last


def classify_point(date_int, start, end):
    """건별 기준일 하나를 분류합니다. 시작일·종료일은 모두 포함합니다."""
    if date_int is None:
        return NO_DATE
    if date_int < start:
        return BEFORE_PERIOD
    if date_int > end:
        return AFTER_PERIOD
    return IN_PERIOD


def classify_span(span_start, span_end, start, end):
    """집계자료의 기간을 분류합니다. 일부만 겹치면 분할불가로 봅니다."""
    if span_start is None and span_end is None:
        return NO_DATE
    lo = span_start if span_start is not None else span_end
    hi = span_end if span_end is not None else span_start
    if lo > hi:
        lo, hi = hi, lo
    if hi < start:
        return BEFORE_PERIOD
    if lo > end:
        return AFTER_PERIOD
    if lo >= start and hi <= end:
        return IN_PERIOD
    return STRADDLE


# ── 집계 재계산 ──────────────────────────────────────────────────

def _sum_by_currency(items, currency_key="currency", amount_key="amount"):
    totals = {}
    for it in items:
        cur = str(it.get(currency_key, "") or "").strip().upper()
        if not cur:
            continue
        totals[cur] = round(totals.get(cur, 0.0) + float(it.get(amount_key, 0) or 0), 2)
    return totals


def _date_bounds(items, date_key="date"):
    values = [d for d in (date_to_int(it.get(date_key)) for it in items) if d]
    return (min(values), max(values)) if values else (None, None)


def _apply_bounds(result, items, start, end, date_key="date"):
    """남은 거래 범위로 문서 기간을 좁힙니다.

    환율 조회 구간과 파일명 라벨이 이 값을 쓰므로, 걸러낸 뒤의 실제 범위여야
    필요 없는 날짜의 환율을 받아오지 않습니다.
    """
    lo, hi = _date_bounds(items, date_key)
    if lo is None:
        doc_lo = date_to_int(result.get("period_start"))
        doc_hi = date_to_int(result.get("period_end"))
        lo = max(doc_lo, start) if doc_lo else start
        hi = min(doc_hi, end) if doc_hi else end
    result["period_start"] = int_to_date(lo)
    result["period_end"] = int_to_date(hi)


# ── 제외 기록 ────────────────────────────────────────────────────

def _record(platform, category, item, label="", date_value=None, currency=""):
    return {
        "platform": platform,
        "category": category,
        "date": str(date_value if date_value is not None else item.get("date", "") or ""),
        # 쇼피처럼 통화가 거래행이 아니라 문서에 있는 플랫폼은 currency 를 넘겨받습니다.
        "currency": str(item.get("currency") or currency or "").upper(),
        "amount": float(item.get("amount", 0) or 0),
        "qty": int(item.get("qty", 0) or 0),
        "label": label or str(
            item.get("order_number") or item.get("order_name") or item.get("order_id")
            or item.get("tracking_no") or ""
        ),
        "source": str(item.get("source_file", "") or ""),
    }


class _Collector:
    def __init__(self):
        self.excluded = []
        self.blocked = []
        self.warnings = []

    def drop(self, platform, category, item, label="", date_value=None, currency=""):
        self.excluded.append(_record(platform, category, item, label, date_value, currency))

    def block(self, platform, label, span_start, span_end):
        self.blocked.append({
            "platform": platform,
            "label": label,
            "span": f"{int_to_date(span_start) or '?'} ~ {int_to_date(span_end) or '?'}",
        })


# ── 플랫폼별 필터 ────────────────────────────────────────────────

def _filter_pointwise(results, platform, start, end, col, item_key="items"):
    """건별 기준일이 있는 자료(쇼피·Joom·쇼피파이)를 거래 단위로 거릅니다."""
    kept_results = []
    for res in results or []:
        out = deepcopy(res)
        # 쇼피는 통화가 문서 단위라 거래행에 currency 키가 없습니다.
        doc_currency = str(res.get("currency", "") or "")
        kept = []
        for it in res.get(item_key, []):
            category = classify_point(date_to_int(it.get("date")), start, end)
            if category == IN_PERIOD:
                kept.append(deepcopy(it))
            else:
                col.drop(platform, category, it, currency=doc_currency)
        if not kept:
            continue
        out[item_key] = kept
        _apply_bounds(out, kept, start, end)
        kept_results.append(out)
    return kept_results


def _filter_shopee(shopee_results, start, end, col):
    kept_results = _filter_pointwise(shopee_results, "쇼피", start, end, col,
                                     item_key="transactions")
    for res in kept_results:
        txs = res["transactions"]
        res["total_qty"] = sum(int(t.get("qty", 0) or 0) for t in txs)
        res["total_amount"] = round(sum(float(t.get("amount", 0) or 0) for t in txs), 2)
    return kept_results


def _filter_joom(joom_results, start, end, col):
    kept_results = _filter_pointwise(joom_results, "Joom", start, end, col)
    for res in kept_results:
        res["total_by_currency"] = _sum_by_currency(res["items"])
        # declared_total 은 원본 PDF 표기 합계이므로 그대로 둡니다. 기간 필터로 일부를
        # 뺀 뒤에는 건별 합산과 다른 것이 정상이라, 합계 불일치 경고를 끕니다.
        res["period_filtered"] = True
        res["total_mismatch"] = False
    return kept_results


def _filter_shopify(shopify_results, start, end, col):
    kept_results = _filter_pointwise(shopify_results, "쇼피파이", start, end, col)
    for res in kept_results:
        items = res["items"]
        res["total_by_currency"] = _sum_by_currency(items)
        res["row_count"] = len(items)
        res["currencies"] = sorted({it.get("currency") for it in items if it.get("currency")})
        # 원본 시트의 반영 표시도 함께 맞춥니다.
        kept_rows = {it.get("row_index") for it in items}
        for flag in res.get("row_flags", []):
            if flag.get("counted") and flag.get("index") not in kept_rows:
                flag["counted"] = False
        res["period_filtered"] = True
    return kept_results


def _filter_lazada(lazada_result, start, end, col):
    """라자다는 주문 Excel(건별 날짜)과 소포수령증 PDF(문서 기간)가 섞일 수 있습니다."""
    if not lazada_result:
        return None
    out = deepcopy(lazada_result)
    dated, undated = [], []
    for it in lazada_result.get("items", []):
        if date_to_int(it.get("date") or it.get("delivered_date")):
            dated.append(it)
        else:
            undated.append(it)

    kept = []
    for it in dated:
        category = classify_point(date_to_int(it.get("date") or it.get("delivered_date")),
                                  start, end)
        if category == IN_PERIOD:
            kept.append(deepcopy(it))
        else:
            col.drop("라자다", category, it,
                     date_value=it.get("date") or it.get("delivered_date"))

    # 건별 날짜가 없는 PDF분은 문서 기간으로 한 덩어리를 판정합니다.
    if undated:
        doc_start = date_to_int(lazada_result.get("period_start"))
        doc_end = date_to_int(lazada_result.get("period_end"))
        category = classify_span(doc_start, doc_end, start, end)
        if category == IN_PERIOD:
            kept.extend(deepcopy(it) for it in undated)
        elif category == STRADDLE:
            label = ", ".join(lazada_result.get("source_files") or []) or "소포수령증 PDF"
            col.block("라자다", f"{label} ({len(undated)}건)", doc_start, doc_end)
        else:
            for it in undated:
                col.drop("라자다", category, it,
                         date_value=lazada_result.get("period_end"))

    if not kept:
        return None
    out["items"] = kept
    out["row_count"] = len(kept)
    out["total_amount_by_currency"] = _sum_by_currency(kept)
    out["total_qty"] = sum(int(it.get("qty", 0) or 0) for it in kept)
    out["currencies"] = sorted({it.get("currency") for it in kept if it.get("currency")})
    _apply_bounds(out, kept, start, end)
    return out


def _filter_ebay(ebay_results, start, end, col):
    """이베이/린코스는 발행월 단위 집계라 월이 경계를 걸치면 나눌 수 없습니다."""
    kept_results = []
    for res in ebay_results or []:
        out = deepcopy(res)
        kept = []
        for it in res.get("items", []):
            span_start, span_end = month_span(it.get("month"))
            if span_start is None:
                span_start = span_end = date_to_int(it.get("date"))
            category = classify_span(span_start, span_end, start, end)
            if category == IN_PERIOD:
                kept.append(deepcopy(it))
            elif category == STRADDLE:
                col.block("이베이", f"{it.get('month') or it.get('date') or '?'}월분",
                          span_start, span_end)
            else:
                col.drop("이베이", category, it, date_value=it.get("date"))
        if not kept:
            continue
        out["items"] = kept
        kept_months = {it.get("month") for it in kept}
        if out.get("summary_items"):
            out["summary_items"] = [s for s in out["summary_items"]
                                    if s.get("month") in kept_months or not s.get("month")]
        _apply_bounds(out, kept, start, end)
        kept_results.append(out)
    return kept_results


def _filter_qoo10(qoo10_result, start, end, col):
    """큐텐은 거래기간 단위이고 신고 기준일은 반기말입니다."""
    if not qoo10_result:
        return None
    entries = qoo10_result.get("entries") or []
    if not entries:
        entries = [{
            "period_start": qoo10_result.get("period_start", ""),
            "period_end": qoo10_result.get("period_end", ""),
            "tracking_no": qoo10_result.get("tracking_no", ""),
            "qty": qoo10_result.get("qty", 0),
            "amount": qoo10_result.get("amount", 0),
            "write_date": qoo10_result.get("write_date", ""),
        }]

    kept = []
    for entry in entries:
        span_start = date_to_int(entry.get("period_start"))
        span_end = date_to_int(entry.get("period_end"))
        category = classify_span(span_start, span_end, start, end)
        label = f"{int_to_date(span_start) or '?'} ~ {int_to_date(span_end) or '?'}"
        if category == IN_PERIOD:
            kept.append(deepcopy(entry))
            # 선적일자로 나가는 반기말이 신고기간 밖이면 그대로 두되 알립니다.
            half_end = _half_year_end(span_end or span_start)
            if half_end and not (start <= half_end <= end):
                col.warnings.append(
                    f"큐텐 {label} 자료의 신고 기준일(반기말 {int_to_date(half_end)})이 "
                    "선택한 신고기간 밖입니다. 신고기간이 반기와 맞는지 확인해 주세요."
                )
        elif category == STRADDLE:
            col.block("큐텐", label, span_start, span_end)
        else:
            col.drop("큐텐", category, {**entry, "currency": "JPY"},
                     label=entry.get("tracking_no", ""),
                     date_value=entry.get("period_end"))

    if not kept:
        return None
    out = deepcopy(qoo10_result)
    out["entries"] = kept
    out["qty"] = sum(int(e.get("qty", 0) or 0) for e in kept)
    out["amount"] = round(sum(float(e.get("amount", 0) or 0) for e in kept), 2)
    out["tracking_no"] = kept[0].get("tracking_no", "")
    out["period_start"] = min((e.get("period_start") for e in kept if e.get("period_start")),
                              default=out.get("period_start", ""))
    out["period_end"] = max((e.get("period_end") for e in kept if e.get("period_end")),
                            default=out.get("period_end", ""))
    return out


def _half_year_end(date_int):
    if not date_int:
        return None
    year, month = date_int // 10000, (date_int // 100) % 100
    return year * 10000 + (630 if month <= 6 else 1231)


# ── 진입점 ───────────────────────────────────────────────────────

def apply_reporting_period(shopee_results=None, lazada_result=None, qoo10_result=None,
                           ebay_results=None, joom_results=None, shopify_results=None,
                           start=None, end=None):
    """신고기간 해당분만 남긴 결과와 분류 내역을 돌려줍니다.

    start/end 가 비어 있으면 필터를 적용하지 않고 원본을 그대로 돌려줍니다.
    분할할 수 없는 집계자료가 경계를 걸치면 `PeriodStraddleError` 를 냅니다.
    """
    start_int = date_to_int(start)
    end_int = date_to_int(end)
    passthrough = {
        "shopee_results": shopee_results or [],
        "lazada_result": lazada_result,
        "qoo10_result": qoo10_result,
        "ebay_results": ebay_results or [],
        "joom_results": joom_results or [],
        "shopify_results": shopify_results or [],
    }
    if start_int is None or end_int is None:
        return {**passthrough, "report": {"applied": False, "excluded": [],
                                          "blocked": [], "warnings": [], "summary": {}}}
    if start_int > end_int:
        raise ValueError("신고기간 시작일이 종료일보다 늦습니다.")

    col = _Collector()
    filtered = {
        "shopee_results": _filter_shopee(shopee_results, start_int, end_int, col),
        "lazada_result": _filter_lazada(lazada_result, start_int, end_int, col),
        "qoo10_result": _filter_qoo10(qoo10_result, start_int, end_int, col),
        "ebay_results": _filter_ebay(ebay_results, start_int, end_int, col),
        "joom_results": _filter_joom(joom_results, start_int, end_int, col),
        "shopify_results": _filter_shopify(shopify_results, start_int, end_int, col),
    }
    if col.blocked:
        raise PeriodStraddleError(col.blocked)

    return {**filtered, "report": {
        "applied": True,
        "start": int_to_date(start_int),
        "end": int_to_date(end_int),
        "excluded": col.excluded,
        "blocked": [],
        "warnings": col.warnings,
        "summary": summarize_exclusions(col.excluded),
    }}


def summarize_exclusions(excluded):
    """제외 내역을 플랫폼·사유별 건수와 통화별 외화금액으로 묶습니다."""
    summary = {}
    for rec in excluded or []:
        key = (rec["platform"], rec["category"])
        bucket = summary.setdefault(key, {"count": 0, "qty": 0, "by_currency": {}})
        bucket["count"] += 1
        bucket["qty"] += int(rec.get("qty", 0) or 0)
        cur = rec.get("currency") or "-"
        bucket["by_currency"][cur] = round(
            bucket["by_currency"].get(cur, 0.0) + float(rec.get("amount", 0) or 0), 2)
    return summary


def format_exclusion_lines(report):
    """처리 로그에 남길 사람이 읽는 형태의 요약 줄을 만듭니다."""
    if not (report or {}).get("applied"):
        return []
    lines = [f"신고기간 {report['start']} ~ {report['end']} 적용"]
    summary = report.get("summary") or {}
    if not summary:
        lines.append("  · 기간 밖 거래 없음 (전 건 신고 해당분)")
    for (platform, category), bucket in sorted(summary.items()):
        amounts = ", ".join(f"{v:,.2f} {c}" for c, v in sorted(bucket["by_currency"].items()))
        lines.append(f"  · {platform} {category} 제외 {bucket['count']:,}건"
                     + (f" / {amounts}" if amounts else ""))
    lines.extend(f"  · {w}" for w in report.get("warnings") or [])
    return lines


def period_presets(year):
    """화면에서 고를 신고기간 프리셋을 만듭니다 (부가세 신고 단위)."""
    return {
        f"{year}년 상반기 (1~6월)": (f"{year}-01-01", f"{year}-06-30"),
        f"{year}년 하반기 (7~12월)": (f"{year}-07-01", f"{year}-12-31"),
        f"{year}년 1기 예정 (1~3월)": (f"{year}-01-01", f"{year}-03-31"),
        f"{year}년 2기 예정 (7~9월)": (f"{year}-07-01", f"{year}-09-30"),
    }
