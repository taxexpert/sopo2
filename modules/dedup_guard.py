# -*- coding: utf-8 -*-
"""중복·충돌·사업자 혼합 검사 — 확정 중복만 자동 제외하고 나머지는 차단하거나 경고합니다.

'같음'을 4단계로 구분합니다 (검증/해결방안.md §4, docs/POLICY.md).

  ① 파일 바이트 동일(SHA-256)            → 확정 중복. 파싱 전에 제외 (app/process에서 처리)
  ② 강한 거래키 동일 + 내용도 전부 동일   → 확정 중복. 자동 제외하고 로그에 남김
  ③ 강한 거래키 동일 + 내용이 다름        → 변경 충돌. 자동 선택하지 않고 생성 중단
  ④ 강한 키가 없는 집계자료               → 자동 제외 금지. 겹침 의심 경고만

플랫폼별 강한 거래키 (샘플 실측으로 유일성 확인):

  쇼피      (통화, 운송장번호)        — 13개 PDF 통합 893건 전부 유일
  라자다    (통화, orderItemId)       — 172건 전부 유일.
            ※ 주문번호·운송장은 키로 쓰지 않는다: 같은 주문의 상품행 분할(주문번호 29종
              반복)과 합포장(운송장 25종 반복)이 정상 데이터다 (v54 금지 회귀 8)
  쇼피파이  (스토어, 주문명 Name)     — 스토어별 유일. 상태만 바뀐 재수출(paid→refunded)은
                                       내용 불일치로 ③ 충돌이 되어 자동 삭제되지 않는다
  Joom      (order_id, 없으면 운송장) — 전건 유일
  이베이    강한 키 없음(월 집계)     — 같은 (발행월, 통화)가 여러 문서에 있으면 경고만
  큐텐      강한 키 없음(수동 목록)   — 동일 내용 입력이 2건 이상이면 경고만
"""

from __future__ import annotations

import re
from copy import deepcopy

from .reporting_period import sum_by_currency


class DuplicateConflictError(RuntimeError):
    """같은 거래키인데 내용이 다른 자료가 있어 생성을 중단합니다."""

    def __init__(self, conflicts):
        self.conflicts = list(conflicts)
        lines = []
        for c in self.conflicts[:10]:
            lines.append(
                f"  · {c['platform']} {c['label']} — 기존 {c['first']} ↔ 새 자료 {c['second']}"
            )
        if len(self.conflicts) > 10:
            lines.append(f"  · … 외 {len(self.conflicts) - 10}건")
        super().__init__(
            "같은 거래인데 금액·날짜·상태가 다른 자료가 함께 올라와 생성을 중단했습니다.\n"
            + "\n".join(lines)
            + "\n어느 쪽이 맞는지 자동으로 고를 수 없습니다. 최신 자료 하나만 남기고 다시 올려주세요."
        )


class MixedSubmitterError(RuntimeError):
    """서로 다른 사업자의 자료가 섞여 있어 생성을 중단합니다."""

    def __init__(self, submitters):
        self.submitters = submitters
        lines = [
            f"  · {info['name'] or '(상호 미확인)'} ({_format_biz(biz)}) — {', '.join(sorted(info['platforms']))}"
            for biz, info in submitters.items()
        ]
        super().__init__(
            "서로 다른 사업자의 자료가 섞여 있어 생성을 중단했습니다.\n"
            + "\n".join(lines)
            + "\n한 번의 생성에는 한 사업자의 자료만 올려주세요."
        )


def _format_biz(digits):
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}" if len(digits) == 10 else digits


# ── 정규화 ───────────────────────────────────────────────────────

def _text(v):
    return str(v or "").strip()


def _amount(v):
    return round(float(v or 0), 2)


def _qty(v):
    return int(v or 0)


def _date8(v):
    d = re.sub(r"\D", "", _text(v))[:8]
    return d if len(d) == 8 else _text(v)


# ── 공통 중복 검사 루프 ──────────────────────────────────────────

def _dedup_items(items, key_fn, content_fn, label_fn, platform, col, seen=None,
                 doc_currency=""):
    """②확정 중복은 버리고 ③내용이 다른 충돌은 기록합니다. 키가 비면 건드리지 않습니다."""
    seen = seen if seen is not None else {}
    kept = []
    for it in items:
        key = key_fn(it)
        if key is None or any(not part for part in key if isinstance(part, str)):
            kept.append(it)          # 식별할 수 없는 행은 자동 제외 금지
            continue
        content = content_fn(it)
        if key not in seen:
            seen[key] = content
            kept.append(it)
        elif seen[key] == content:
            col["dropped"].append({
                "platform": platform, "label": label_fn(it),
                # 쇼피는 통화가 거래행이 아니라 문서에 있으므로 doc_currency 로 보완
                "currency": _text(it.get("currency")) or _text(doc_currency),
                "amount": _amount(it.get("amount")),
                "content": content,
            })
        else:
            col["conflicts"].append({
                "platform": platform, "label": label_fn(it),
                "first": _describe(seen[key]), "second": _describe(content),
            })
    return kept, seen


def _describe(content):
    return " / ".join(str(v) for v in content if str(v))


# ── 플랫폼별 키·내용 정의 ────────────────────────────────────────

def _dedup_shopee(shopee_results, col):
    seen = {}
    kept_results = []
    for res in shopee_results or []:
        cur = _text(res.get("currency"))
        kept, seen = _dedup_items(
            res.get("transactions", []),
            key_fn=lambda t, c=cur: (c, _text(t.get("tracking_no")).upper()),
            content_fn=lambda t: (_date8(t.get("date")), _qty(t.get("qty")), _amount(t.get("amount"))),
            label_fn=lambda t: _text(t.get("tracking_no")),
            platform="쇼피", col=col, seen=seen, doc_currency=cur,
        )
        if not kept:
            continue                 # 전부 중복인 문서는 시트를 만들지 않음 (v54 §3.3)
        out = deepcopy(res)
        out["transactions"] = deepcopy(kept)
        out["total_qty"] = sum(_qty(t.get("qty")) for t in kept)
        out["total_amount"] = round(sum(_amount(t.get("amount")) for t in kept), 2)
        kept_results.append(out)
    return kept_results


def _dedup_lazada(lazada_result, col):
    if not lazada_result:
        return None
    kept, _ = _dedup_items(
        lazada_result.get("items", []),
        # 주문번호·운송장은 상품행 분할·합포장으로 정상 반복되므로 키로 쓰지 않는다
        # (v54 금지 회귀 8, 샘플 실측: orderItemId 만 전건 유일)
        key_fn=lambda i: (_text(i.get("currency")), _text(i.get("order_item_id"))),
        content_fn=lambda i: (_amount(i.get("amount")), _qty(i.get("qty")),
                              _date8(i.get("date") or i.get("delivered_date")),
                              _text(i.get("status")), _text(i.get("tracking_no"))),
        label_fn=lambda i: f"주문상품 {_text(i.get('order_item_id'))}",
        platform="라자다", col=col,
    )
    if not kept:
        return None
    out = deepcopy(lazada_result)
    out["items"] = deepcopy(kept)
    out["row_count"] = len(kept)
    out["total_qty"] = sum(_qty(i.get("qty")) for i in kept)
    out["total_amount_by_currency"] = sum_by_currency(kept)
    out["currencies"] = sorted({i.get("currency") for i in kept if i.get("currency")})
    return out


def _dedup_shopify(shopify_results, col):
    seen = {}
    kept_results = []
    for res in shopify_results or []:
        store = _text(res.get("store"))
        kept, seen = _dedup_items(
            res.get("items", []),
            key_fn=lambda i, s=store: (s, _text(i.get("order_name"))),
            # financial_status 를 내용에 포함해, 상태만 바뀐 재수출(paid→refunded)은
            # 확정 중복이 아니라 ③변경 충돌로 잡혀 생성이 중단된다 (검증기준 T05)
            content_fn=lambda i: (_amount(i.get("amount")), _text(i.get("currency")),
                                  _date8(i.get("date")), _text(i.get("financial_status")),
                                  _qty(i.get("qty"))),
            label_fn=lambda i, s=store: f"{s} {_text(i.get('order_name'))}",
            platform="쇼피파이", col=col, seen=seen,
        )
        if not kept:
            continue
        out = deepcopy(res)
        out["items"] = deepcopy(kept)
        out["total_by_currency"] = sum_by_currency(kept)
        out["row_count"] = len(kept)
        out["currencies"] = sorted({i.get("currency") for i in kept if i.get("currency")})
        kept_rows = {i.get("row_index") for i in kept}
        for flag in out.get("row_flags", []):
            if flag.get("counted") and flag.get("index") not in kept_rows:
                flag["counted"] = False
        kept_results.append(out)
    return kept_results


def _dedup_joom(joom_results, col):
    seen = {}
    kept_results = []
    for res in joom_results or []:
        kept, seen = _dedup_items(
            res.get("items", []),
            key_fn=lambda i: ("joom", _text(i.get("order_id")) or _text(i.get("tracking_no")).upper()),
            content_fn=lambda i: (_amount(i.get("amount")), _text(i.get("currency")),
                                  _date8(i.get("date")), _qty(i.get("qty"))),
            label_fn=lambda i: _text(i.get("order_id")) or _text(i.get("tracking_no")),
            platform="Joom", col=col, seen=seen,
        )
        if not kept:
            continue
        out = deepcopy(res)
        dropped_here = len(res.get("items", [])) - len(kept)
        out["items"] = deepcopy(kept)
        out["total_by_currency"] = sum_by_currency(kept)
        if dropped_here:
            # declared_total 은 원본 PDF 표기라 중복 제외 후에는 건별 합산과 달라도 정상
            out["total_mismatch"] = False
            out["dedup_filtered"] = True
        kept_results.append(out)
    return kept_results


def _warn_ebay_overlap(ebay_results, col):
    """이베이 월집계는 강한 키가 없어 자동 제외하지 않고 겹침 의심만 알립니다."""
    seen = {}
    for idx, res in enumerate(ebay_results or []):
        for it in res.get("items", []):
            key = (_text(it.get("month")), _text(it.get("currency")))
            if all(key) and key in seen and seen[key] != idx:
                col["warnings"].append(
                    f"이베이 {key[0]} ({key[1]}) 월 집계가 여러 파일에 있습니다. "
                    "같은 월분이 중복 반영되지 않았는지 확인해 주세요."
                )
            seen.setdefault(key, idx)


def _warn_qoo10_duplicates(qoo10_result, col):
    """큐텐 STEP 2 목록은 수동 입력이라 자동 삭제하지 않고 동일 내용만 알립니다."""
    if not qoo10_result:
        return
    seen = set()
    for e in qoo10_result.get("entries") or []:
        key = (_date8(e.get("period_start")), _date8(e.get("period_end")),
               _text(e.get("tracking_no")), _amount(e.get("amount")), _qty(e.get("qty")))
        if key in seen:
            col["warnings"].append(
                f"큐텐 입력 목록에 동일한 건이 2번 있습니다: 발송번호 {key[2] or '(없음)'} / "
                f"{key[3]:,.0f} JPY. STEP 2에서 중복 입력인지 확인해 주세요."
            )
        seen.add(key)


# ── 진입점 ───────────────────────────────────────────────────────

def dedup_transactions(shopee_results=None, lazada_result=None, qoo10_result=None,
                       ebay_results=None, joom_results=None, shopify_results=None):
    """확정 중복(②)을 걷어낸 결과와 검사 내역을 돌려줍니다.

    변경 충돌(③)이 있으면 `DuplicateConflictError` 로 생성을 중단합니다.
    원본 dict 는 건드리지 않습니다.
    """
    col = {"dropped": [], "conflicts": [], "warnings": []}
    cleaned = {
        "shopee_results": _dedup_shopee(shopee_results, col),
        "lazada_result": _dedup_lazada(lazada_result, col),
        "qoo10_result": deepcopy(qoo10_result) if qoo10_result else None,
        "ebay_results": [deepcopy(r) for r in (ebay_results or [])],
        "joom_results": _dedup_joom(joom_results, col),
        "shopify_results": _dedup_shopify(shopify_results, col),
    }
    _warn_ebay_overlap(ebay_results, col)
    _warn_qoo10_duplicates(qoo10_result, col)

    if col["conflicts"]:
        raise DuplicateConflictError(col["conflicts"])

    return {**cleaned, "report": {
        "dropped": col["dropped"],
        "warnings": col["warnings"],
    }}


def format_dedup_lines(report):
    """처리 로그에 남길 요약 줄을 만듭니다."""
    lines = []
    dropped = (report or {}).get("dropped") or []
    if dropped:
        by_platform = {}
        for d in dropped:
            b = by_platform.setdefault(d["platform"], {"count": 0, "by_currency": {}})
            b["count"] += 1
            cur = d["currency"] or "-"
            b["by_currency"][cur] = round(b["by_currency"].get(cur, 0.0) + d["amount"], 2)
        for platform, b in sorted(by_platform.items()):
            amounts = ", ".join(f"{v:,.2f} {c}" for c, v in sorted(b["by_currency"].items()))
            lines.append(f"[중복] {platform} 동일 거래 {b['count']:,}건 제외 ({amounts})")
    for w in (report or {}).get("warnings") or []:
        lines.append(f"[WARN] {w}")
    return lines


def check_submitter_mix(shopee_results=None, lazada_result=None, qoo10_result=None,
                        ebay_results=None, joom_results=None, shopify_results=None):
    """자료마다 읽힌 제출자 사업자번호가 2개 이상이면 생성을 중단합니다.

    쇼피파이 orders 처럼 제출자 정보가 없는 자료는 판정에 쓰지 않고,
    미확인 제출자는 기존 정책대로 중립 문구로 계속 생성합니다 (v54 §3.1).
    반환값은 경고 문자열 목록입니다.
    """
    sources = []
    for res in shopee_results or []:
        sources.append(("쇼피", res.get("submitter") or {}))
    if lazada_result:
        sources.append(("라자다", lazada_result.get("submitter") or {}))
    if qoo10_result:
        sources.append(("큐텐", qoo10_result.get("submitter") or {}))
    for res in ebay_results or []:
        sources.append(("이베이", res.get("submitter") or {}))
    for res in joom_results or []:
        sources.append(("Joom", res.get("submitter") or {}))
    # 쇼피파이 orders 에는 사업자 정보가 없어 혼합 판정에 사용하지 않습니다.

    by_biz = {}
    names_without_biz = set()
    for platform, sub in sources:
        name = _text(sub.get("name"))
        biz = re.sub(r"\D", "", _text(sub.get("biz_no")))
        if biz:
            info = by_biz.setdefault(biz, {"name": name, "platforms": set()})
            info["platforms"].add(platform)
            if name and not info["name"]:
                info["name"] = name
        elif name:
            names_without_biz.add(name)

    if len(by_biz) > 1:
        raise MixedSubmitterError(by_biz)

    warnings = []
    known_names = {info["name"] for info in by_biz.values() if info["name"]}
    strange = names_without_biz - known_names
    if known_names and strange:
        warnings.append(
            "사업자번호 없이 상호만 다른 자료가 섞여 있습니다: "
            + ", ".join(sorted(strange))
            + " — 같은 사업자인지 확인해 주세요."
        )
    return warnings
