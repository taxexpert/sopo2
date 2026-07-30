# -*- coding: utf-8 -*-
"""
소포수령증 자동화 웹앱 — v53 Joom · 쇼피파이 · 린코스 지원
실행: streamlit run app.py
"""

from __future__ import annotations

import contextlib
import html
import io
import os
import hashlib
import re
import sys
import tempfile
import time
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from modules.pdf_parser import parse_pdf, detect_pdf_type
from modules.lazada_order_parser import (
    is_lazada_order_excel, parse_lazada_order_excel, merge_lazada_results,
)
from modules.shopify_parser import (
    is_shopify_orders_file, parse_shopify_orders, merge_shopify_results,
)
from modules.excel_writer import generate_excel, period_labels
from modules.exchange_rate import (
    fetch_all_currencies_for_period,
    fetch_monthly_avg_currencies_for_period,
    merge_monthly_rates,
    RATE_LOOKBACK_DAYS,
)
from modules.extra_docs import (
    build_declaration_rows,
    company_name_from_results,
    create_export_performance,
    create_zero_rate_attachments,
    safe_filename,
)
from modules.reporting_period import (
    PeriodStraddleError,
    apply_reporting_period,
    format_exclusion_lines,
    period_presets,
)
from modules.dedup_guard import (
    DuplicateConflictError,
    MixedSubmitterError,
    check_submitter_mix,
    dedup_transactions,
    format_dedup_lines,
)

CURRENCIES = ["MYR", "PHP", "SGD", "THB", "TWD", "VND", "IDR", "JPY", "BRL", "MXN", "USD", "EUR", "GBP", "CAD", "AUD"]

# ── 선택적 로그인: Streamlit secrets [auth]가 있을 때만 사용 ────────────
ALLOWED_EMAILS = [
    "guwjd2298@gmail.com",
    "help@taxexpert.kr",
    "m0120@taxexpert.kr",
    "m0227@taxexpert.kr",
    "m0125@taxexpert.kr",
    "ayoung9976@gmail.com",
    "m0429@taxexpert.kr",
    "m0607@taxexpert.kr",
    "m1007@taxexpert.kr",
    "m1211@taxexpert.kr",
    "m1225@taxexpert.kr",
    "m1018@taxexpert.kr",
]

st.set_page_config(page_title="소포수령증 자동화", page_icon="📦", layout="centered")

# 로그인 화면과 본 화면이 같은 스타일을 쓰도록 인증 처리보다 먼저 주입합니다.
# 화면 언어는 세무 서식을 따릅니다. 구획은 괘선으로 나누고, 강조색은 번호·기본동작에만 씁니다.
st.markdown(
    """
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.css');

/* 화면 글자 크기의 기준값입니다. Streamlit 내부 크기와 아래 지정이 모두 rem이라
   본문·설명글·버튼·표가 함께 비례합니다. 브라우저 기본값 16px에는 규칙이 없어 이 한 줄로 잡힙니다.
   config.toml의 theme.baseFontSize로 하면 안 됩니다. [theme] 섹션을 선언하는 순간
   Streamlit이 브라우저의 밝음/어둠 설정을 따르지 않고 밝은 테마로 고정돼,
   아래 prefers-color-scheme 규칙과 어긋나 흰 배경에 밝은 글자가 됩니다. */
html { font-size: 17px; }

/* 색은 엑스퍼트 로고에서 가져온 두 가지뿐입니다.
   브랜드 블루가 강조를 맡고 하늘색이 보조로 받칩니다. 중성색도 회색 대신 파랑 기미를 둡니다. */
:root {
    --ink:#0f1d30; --ink-soft:#55627a; --line:#d8e4f5;
    /* 하늘색은 로고에선 파란 바탕 위라 밝아도 읽히지만, 흰 바탕에선 한 단계 내려야 보입니다. */
    --accent:#1069f5; --sky:#0e9ad4; --surface:#f2f7fe; --surface-tint:#e6f0fd;
    /* 채워진 버튼은 배경·글자를 따로 둡니다. 강조색을 그대로 배경에 쓰면 어두운 테마에서 대비가 깨집니다. */
    --accent-solid:#1069f5; --accent-hover:#0b57d0; --on-accent:#ffffff;
}
@media (prefers-color-scheme: dark) {
    :root {
        --ink:#e6edf8; --ink-soft:#93a1ba; --line:#27354a;
        --accent:#5c9dff; --sky:#7bd8f7; --surface:#131c28; --surface-tint:#182437;
        --accent-solid:#5c9dff; --accent-hover:#8ab8ff; --on-accent:#08111f;
    }
}

/* 한글 본문용 서체. CDN이 막히면 OS 기본 한글 서체로 되돌아갑니다.
   Streamlit 아이콘(Material Symbols)은 제외해야 글리프가 글자로 깨지지 않습니다. */
.stApp, .stApp p, .stApp li, .stApp td, .stApp th, .stApp label,
.stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp button, .stApp input, .stApp textarea, .stApp select {
    font-family:'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont,
        'Apple SD Gothic Neo', 'Malgun Gothic', 'Segoe UI', Roboto, sans-serif;
}
.stApp [class*="material-symbols"], .stApp [data-testid="stIconMaterial"] {
    font-family:'Material Symbols Rounded' !important;
}

/* ── 문서 머리 ── */
.doc-head { margin:0.25rem 0 1.6rem; }
.doc-head__brand { display:flex; align-items:center; gap:0.7rem; }
/* 로고의 2톤 구성을 그대로 따릅니다. 몸통은 브랜드 블루, 봉함 띠는 하늘색. */
.doc-head__mark { flex:none; width:34px; height:34px; color:var(--accent); }
.doc-head__mark .mark-tape { stroke:var(--sky); }
.doc-head h1.doc-head__title { font-size:1.85rem; font-weight:700; letter-spacing:-0.015em;
    line-height:1.2; color:var(--ink); margin:0; padding:0; }
/* 폭 제한을 두지 않아야 아래 구획 상자들과 오른쪽 끝선이 맞습니다. */
.doc-head p.doc-head__sub { font-size:0.95rem; line-height:1.6; color:var(--ink-soft);
    margin:0.6rem 0 0; }

/* ── 단계 구획 ── */
.step { display:flex; align-items:baseline; gap:0.7rem;
    border-top:1px solid var(--line); padding-top:1.9rem; margin-bottom:0.7rem; }
.step--lead { border-top:0; padding-top:0; }
/* 번호는 제목과 같은 크기로 맞춰야 baseline 정렬에서 위끝도 나란해집니다. */
.step__no { flex:none; font-size:1.12rem; font-weight:700; color:var(--accent);
    font-variant-numeric:tabular-nums; line-height:1.3; }
.step__rule { flex:none; width:1px; align-self:stretch; background:var(--line); }
.step h2.step__title { font-size:1.12rem; font-weight:650; letter-spacing:-0.01em;
    line-height:1.3; color:var(--ink); margin:0; padding:0; }

/* ── 안내문 ── */
.note { border:1px solid var(--line); border-radius:6px; background:var(--surface);
    padding:0.7rem 0.9rem; font-size:0.875rem; line-height:1.6; color:var(--ink-soft);
    margin:0.2rem 0 0.6rem; }

/* ── 업로드 파일 목록 ── */
.file-row { display:flex; align-items:center; gap:0.5rem; padding:0.3rem 0; min-width:0; }
.file-row__name { font-size:0.875rem; color:var(--ink); overflow:hidden;
    text-overflow:ellipsis; white-space:nowrap; }
/* 판별에 성공한 파일은 브랜드 블루, 못 한 파일은 앰버로 한눈에 갈립니다. */
.file-row__tag { flex:none; font-size:0.75rem; line-height:1; color:var(--accent);
    background:var(--surface-tint); border:1px solid var(--line); border-radius:4px;
    padding:0.25rem 0.4rem; white-space:nowrap; }
.file-row__tag--unknown { color:#b45309; background:#fdf5e7; border-color:#f0c98a; }

.log-box { background:#0b1020; color:#e7edf8; border-radius:8px; padding:1rem;
    font-family:'JetBrains Mono', Consolas, monospace; font-size:0.85rem;
    max-height:320px; overflow:auto; white-space:pre-wrap; }
div[data-testid="stColumn"] .stButton > button { white-space:nowrap !important; }

/* ── 접기 구획 ── */
/* 두 곳(큐텐재팬 목록·파일명 규칙)이 단계 머리와 같은 괘선색을 쓰도록 맞춥니다.
   글자 크기는 지정하지 않습니다. summary가 inherit이라 기준 크기를 이미 따릅니다. */
.stApp [data-testid="stExpander"] details { border-color:var(--line); border-radius:6px; }
.stApp [data-testid="stExpander"] summary { font-weight:600; color:var(--ink); }

/* 기본 동작 버튼은 Streamlit 기본 빨강 대신 문서 강조색을 씁니다. */
.stApp .stButton button[kind="primary"]:not(:disabled),
.stApp .stFormSubmitButton button[kind="primary"]:not(:disabled),
.stApp button[data-testid="baseButton-primary"]:not(:disabled) {
    background-color:var(--accent-solid); border-color:var(--accent-solid);
    color:var(--on-accent);
}
/* 조건이 갖춰지기 전에는 눌러도 소용없다는 것이 색으로 먼저 보여야 합니다. */
.stApp .stButton button[kind="primary"]:disabled,
.stApp .stFormSubmitButton button[kind="primary"]:disabled,
.stApp button[data-testid="baseButton-primary"]:disabled {
    background-color:var(--surface); border-color:var(--line);
    color:var(--ink-soft); opacity:1;
}
.stApp .stButton button[kind="primary"]:hover:not(:disabled),
.stApp .stFormSubmitButton button[kind="primary"]:hover:not(:disabled) {
    background-color:var(--accent-hover); border-color:var(--accent-hover);
    color:var(--on-accent);
}

/* ── 저장소 바로가기 ── */
/* Streamlit의 링크 색 규칙을 넘어서야 하므로 선택자를 조이고 !important를 함께 씁니다. */
.stApp a.repo-link { position:fixed; right:1.5rem; bottom:3.25rem; z-index:999;
    width:44px; height:44px; border-radius:50%;
    display:inline-flex; align-items:center; justify-content:center;
    border:1px solid var(--line) !important; background:var(--surface);
    color:var(--ink-soft) !important; text-decoration:none !important; }
.stApp a.repo-link:hover { color:var(--ink) !important; border-color:var(--ink-soft) !important; }
.stApp a.repo-link:focus-visible { outline:2px solid var(--accent); outline-offset:2px; }
.stApp a.repo-link svg { width:19px; height:19px; fill:currentColor; }

@media (max-width:640px) {
    .doc-head h1.doc-head__title { font-size:1.55rem; }
    .doc-head__mark { width:30px; height:30px; }
    .stApp a.repo-link { right:1rem; bottom:1rem; }
}
</style>
""",
    unsafe_allow_html=True,
)

# 소포 마크와 저장소 링크는 로그인 화면·본 화면에서 함께 씁니다.
PARCEL_MARK = (
    '<svg class="doc-head__mark" viewBox="0 0 32 32" fill="none" stroke="currentColor" '
    'stroke-width="1.7" stroke-linejoin="round" aria-hidden="true">'
    '<rect x="4.5" y="8.5" width="23" height="19" rx="2"/>'
    '<path class="mark-tape" d="M4.5 14.75H27.5"/>'
    '<path class="mark-tape" d="M16 8.5V27.5"/></svg>'
)
REPO_LINK = (
    '<a class="repo-link" href="https://github.com/taxexpert/sopo2" target="_blank" '
    'rel="noopener noreferrer" title="GitHub 저장소" aria-label="GitHub 저장소 열기">'
    '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 '
    '2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94'
    '-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 '
    '2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36'
    '-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 '
    '2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54'
    '.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"/>'
    '</svg></a>'
)


def doc_head(subtitle: str) -> None:
    """제품명·소포 마크·한 줄 설명으로 구성된 문서 머리를 그립니다."""
    st.markdown(
        f'<header class="doc-head">'
        f'<div class="doc-head__brand">{PARCEL_MARK}'
        f'<h1 class="doc-head__title">소포수령증 자동화</h1></div>'
        f'<p class="doc-head__sub">{subtitle}</p></header>',
        unsafe_allow_html=True,
    )


def step_head(no: int, title: str, lead: bool = False) -> None:
    """번호·세로 괘선·제목으로 이루어진 단계 머리를 그립니다."""
    st.markdown(
        f'<div class="step{" step--lead" if lead else ""}">'
        f'<span class="step__no">{no}</span><span class="step__rule"></span>'
        f'<h2 class="step__title">{title}</h2></div>',
        unsafe_allow_html=True,
    )


def note(text: str) -> None:
    """괘선 상자로 감싼 보조 안내문을 그립니다."""
    st.markdown(f'<div class="note">{text}</div>', unsafe_allow_html=True)

_AUTH_ENABLED = False
try:
    _AUTH_ENABLED = "auth" in st.secrets
except Exception:
    _AUTH_ENABLED = False

if _AUTH_ENABLED:
    try:
        _logged_in = st.user.is_logged_in
    except Exception as e:
        st.error(f"로그인 상태 확인 오류: {type(e).__name__}: {e}")
        st.stop()
    if not _logged_in:
        st.markdown('<div style="height:3rem"></div>', unsafe_allow_html=True)
        doc_head("업무용 계정으로 로그인하면 이어서 진행할 수 있습니다.")
        c1, c2, c3 = st.columns([1, 1.4, 1])
        with c2:
            st.button("Google 계정으로 로그인", type="primary", on_click=st.login, use_container_width=True)
        st.markdown(REPO_LINK, unsafe_allow_html=True)
        st.stop()
    _user_email = st.user.get("email", "")
    if ALLOWED_EMAILS and _user_email not in ALLOWED_EMAILS:
        st.error(f"접근 권한이 없는 계정입니다. ({_user_email})")
        if st.button("로그아웃"):
            st.logout()
        st.stop()
    with st.sidebar:
        st.markdown(f"**{st.user.get('name','') or _user_email}**")
        st.caption(_user_email)
        if st.button("로그아웃"):
            st.logout()

doc_head(
    "수령증 PDF와 주문내역을 올리면 매출집계·영세율첨부서류제출명세서·수출실적명세서를 "
    "매매기준율까지 적용해 한 번에 만듭니다."
)
st.markdown(REPO_LINK, unsafe_allow_html=True)

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "qoo10_entries" not in st.session_state:
    st.session_state.qoo10_entries = []
if "result_files" not in st.session_state:
    st.session_state.result_files = []
if "qoo10_auto_imported_keys" not in st.session_state:
    st.session_state.qoo10_auto_imported_keys = set()
# 큐텐 구획은 기본으로 접혀 있습니다. 담당자가 한 번 손대면 그 세션 동안 열어 둡니다.
if "qoo10_panel_open" not in st.session_state:
    st.session_state.qoo10_panel_open = False

# 파일명으로 구분되지 않는 PDF에만 직접 선택 항목을 표시합니다.
# 일반 화면의 문구와 배치는 v38과 동일하게 유지합니다.
UNKNOWN_PDF_TYPE_OPTIONS = ["쇼피", "라자다", "큐텐재팬", "이베이/린코스", "Joom"]
UNKNOWN_PDF_TYPE_TO_CODE = {
    "쇼피": "shopee",
    "라자다": "lazada",
    "큐텐재팬": "qoo10",
    "이베이/린코스": "ebay",
    "Joom": "joom",
}

FILE_TYPE_LABELS = {
    "shopee": "쇼피", "lazada": "라자다", "lazada_excel": "라자다 주문내역",
    "qoo10": "큐텐재팬", "ebay": "이베이/린코스", "joom": "Joom",
    "shopify": "쇼피파이 주문내역",
    "unknown": "미확인", "unknown_excel": "미확인 Excel", "unknown_csv": "미확인 CSV",
}

# ══════════════════════════════════════════════════════════════════
# STEP 1 — PDF / Excel 업로드
# ══════════════════════════════════════════════════════════════════
step_head(1, "수령증·주문내역 올리기", lead=True)
# 캡션이 한 줄일 때는 캡션이, 두 줄로 감길 때는 버튼이 서로의 중심선에 맞습니다.
# CSS로 margin을 박으면 둘 중 한쪽에서 반드시 어긋나므로 열 정렬 인자를 씁니다.
c_desc, c_reset = st.columns([5, 1], vertical_alignment="center")
c_desc.caption("쇼피·라자다·큐텐재팬·이베이(린코스)·Joom PDF, 라자다 주문내역 Excel, 쇼피파이 orders CSV를 함께 올릴 수 있습니다.")
# 열을 가득 채워야 버튼 오른쪽 끝이 아래 상자들의 끝선과 맞습니다.
if c_reset.button("초기화", use_container_width=True):
    st.session_state.uploader_key += 1
    st.session_state.qoo10_entries = []
    st.session_state.result_files = []
    st.session_state.qoo10_auto_imported_keys = set()
    # 자동으로 펼쳐진 구획은 이때 다시 접힙니다. 담당자가 직접 펼친 경우에는
    # Streamlit이 그 선택을 기억해 계속 열어 두므로 접히지 않습니다.
    st.session_state.qoo10_panel_open = False
    st.rerun()

note("플랫폼·배송업체가 발행한 정해진 서식의 수령증·주문내역만 읽습니다. "
     "직접 만든 집계표나 서식이 다른 자료, 화면을 캡처한 이미지는 인식하지 못합니다. "
     "서식이 다르면 파일 목록에 '미확인'으로 표시됩니다.")

uploaded_files = st.file_uploader(
    "PDF · Excel · CSV 파일 선택",
    type=["pdf", "xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key=f"file_uploader_{st.session_state.uploader_key}",
)

@st.cache_data(show_spinner=False)
def _detect_uploaded_file_type(filename: str, payload: bytes) -> str:
    """PDF는 본문 표식, Excel/CSV는 필수 열로 플랫폼을 자동 판별합니다."""
    suffix = Path(filename).suffix.lower()
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / (Path(filename).name or f"uploaded{suffix}")
        path.write_bytes(payload)
        if suffix == ".csv":
            return "shopify" if is_shopify_orders_file(path) else "unknown_csv"
        if suffix in {".xlsx", ".xlsm"}:
            if is_lazada_order_excel(path):
                return "lazada_excel"
            if is_shopify_orders_file(path):
                return "shopify"
            return "unknown_excel"
        return detect_pdf_type(str(path))


@st.cache_data(show_spinner=False)
def _parse_uploaded_qoo10(filename: str, payload: bytes):
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / (Path(filename).name or "qoo10.pdf")
        path.write_bytes(payload)
        return parse_pdf(str(path), forced_type="qoo10")


uploaded_type_choices = {}
if uploaded_files:
    cols = st.columns(2)
    for i, f in enumerate(uploaded_files):
        payload = f.getvalue()
        ptype = _detect_uploaded_file_type(f.name, payload)
        label = FILE_TYPE_LABELS.get(ptype, "판별 중")
        tag_mod = " file-row__tag--unknown" if ptype.startswith("unknown") else ""
        target_col = cols[i % 2]
        target_col.markdown(
            f'<div class="file-row"><span class="file-row__name" title="{html.escape(f.name, quote=True)}">'
            f'{html.escape(f.name)}</span>'
            f'<span class="file-row__tag{tag_mod}">{html.escape(label)}</span></div>',
            unsafe_allow_html=True,
        )
        if ptype == "unknown":
            selected_label = target_col.selectbox(
                "문서 종류",
                UNKNOWN_PDF_TYPE_OPTIONS,
                key=f"pdf_type_{st.session_state.uploader_key}_{i}_{f.name}",
                label_visibility="collapsed",
            )
            uploaded_type_choices[f.name] = UNKNOWN_PDF_TYPE_TO_CODE[selected_label]
        elif ptype == "unknown_excel":
            target_col.error(
                "인식할 수 없는 Excel입니다. 라자다는 deliveredDate·paidPrice, "
                "쇼피파이는 Name·Financial Status·Fulfilled at·Total 열이 필요합니다."
            )
            uploaded_type_choices[f.name] = ptype
        elif ptype == "unknown_csv":
            target_col.error(
                "쇼피파이 주문내역(orders) CSV가 아닙니다. "
                "transaction·payout 파일이 아닌 orders 파일을 올려주세요."
            )
            uploaded_type_choices[f.name] = ptype
        else:
            uploaded_type_choices[f.name] = ptype

    # 큐텐재팬 PDF는 업로드 즉시 파싱하여 STEP 2 입력 목록에 자동 반영합니다.
    qoo10_uploads = []
    for f in uploaded_files:
        if uploaded_type_choices.get(f.name) != "qoo10":
            continue
        payload = f.getvalue()
        file_key = hashlib.sha256(payload).hexdigest()
        qoo10_uploads.append((f, payload, file_key))

    # 업로더에서 제거된 PDF의 자동 입력 행은 STEP 2에서도 함께 제거합니다.
    current_qoo10_keys = {key for _, _, key in qoo10_uploads}
    st.session_state.qoo10_entries = [
        entry for entry in st.session_state.qoo10_entries
        if not entry.get("_auto_imported") or entry.get("_file_key") in current_qoo10_keys
    ]
    st.session_state.qoo10_auto_imported_keys = set(st.session_state.qoo10_auto_imported_keys) & current_qoo10_keys

    existing_auto_keys = {
        entry.get("_file_key") for entry in st.session_state.qoo10_entries
        if entry.get("_auto_imported") and entry.get("_file_key")
    }
    for f, payload, file_key in qoo10_uploads:
        if file_key in existing_auto_keys or file_key in st.session_state.qoo10_auto_imported_keys:
            continue
        result = _parse_uploaded_qoo10(f.name, payload)
        st.session_state.qoo10_auto_imported_keys.add(file_key)
        if not result:
            continue
        st.session_state.qoo10_entries.append({
            "period_start": result.get("period_start", ""),
            "period_end": result.get("period_end", ""),
            "tracking_no": result.get("tracking_no", ""),
            "qty": int(result.get("qty", 0) or 0),
            "amount": float(result.get("amount", 0) or 0),
            "write_date": result.get("write_date", ""),
            "_source_file": f.name,
            "_file_key": file_key,
            "_auto_imported": True,
            "_submitter": result.get("submitter") or {},
            "_warnings": result.get("refund_warnings") or [],
        })

# ══════════════════════════════════════════════════════════════════
# STEP 2 — 큐텐재팬 정보 입력
# ══════════════════════════════════════════════════════════════════
step_head(2, "큐텐재팬 정보 확인")


def _fmt_date(v: str) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    if len(d) == 8:
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return str(v or "").strip()


def _open_qoo10_panel() -> None:
    """추가 버튼을 누르면 경고나 결과가 접힌 채 묻히지 않도록 구획을 열어 둡니다."""
    st.session_state.qoo10_panel_open = True


# 라벨에 건수를 드러내야 열지 않고도 안에 무엇이 있는지 알 수 있습니다.
# STEP 1의 자동입력이 이 지점보다 먼저 실행되므로, 큐텐 PDF를 올린 그 실행에서 바로 펼쳐집니다.
_qoo10_count = len(st.session_state.qoo10_entries)
_qoo10_label = (
    f"큐텐재팬 입력 목록 ({_qoo10_count}건)" if _qoo10_count
    else "큐텐재팬 입력 목록 (0건) · 큐텐 자료가 없으면 넘어가세요"
)

with st.expander(
    _qoo10_label,
    expanded=bool(_qoo10_count) or st.session_state.qoo10_panel_open,
):
    note("큐텐재팬 PDF를 올리면 아래 목록에 자동으로 채워집니다. 빠진 건은 직접 추가해주세요.")

    with st.form("qoo10_add_form", clear_on_submit=True):
        fp1, fp2 = st.columns(2)
        in_ps = fp1.text_input("거래기간 시작일", placeholder="예: 20260101")
        in_pe = fp2.text_input("거래기간 종료일", placeholder="예: 20260131")
        fc1, fc2, fc3, fc4 = st.columns(4)
        in_amount = fc1.number_input("금액(JPY)", min_value=0, value=0, format="%d")
        in_qty = fc2.number_input("건수", min_value=0, value=0, format="%d")
        in_track = fc3.text_input("발송번호", placeholder="예: K2512244647017")
        in_wdate = fc4.text_input("발행일", placeholder="예: 20260205")
        # 접힌 구획도 내용을 렌더링하지만 inert로 가려지므로, 유효성 경고가 묻히지 않도록
        # 제출 시점에 구획을 열어 둡니다. 목록이 비면 접히는 조건과 맞물리는 지점입니다.
        added = st.form_submit_button(
            "목록에 추가", use_container_width=True, on_click=_open_qoo10_panel
        )

    if added:
        if in_amount > 0 or in_qty > 0 or in_track.strip():
            st.session_state.qoo10_entries.append({
                "period_start": _fmt_date(in_ps),
                "period_end": _fmt_date(in_pe),
                "tracking_no": in_track.strip(),
                "qty": int(in_qty),
                "amount": float(in_amount),
                "write_date": _fmt_date(in_wdate),
            })
            # 라벨의 건수는 이 지점보다 앞에서 계산되므로, 다시 그려야 표와 숫자가 맞습니다.
            st.rerun()
        else:
            st.warning("금액·건수·발송번호 중 하나는 입력해야 합니다.")

    if st.session_state.qoo10_entries:
        visible_cols = ["period_start", "period_end", "tracking_no", "qty", "amount", "write_date"]
        df_show = pd.DataFrame(st.session_state.qoo10_entries)[visible_cols].rename(columns={
            "period_start": "거래기간 시작",
            "period_end": "거래기간 종료",
            "tracking_no": "발송번호",
            "qty": "건수",
            "amount": "금액(JPY)",
            "write_date": "발행일",
        })
        df_show.index = range(1, len(df_show) + 1)
        df_show["금액(JPY)"] = df_show["금액(JPY)"].map(lambda x: f"{int(x):,}")
        df_show["건수"] = df_show["건수"].map(lambda x: f"{int(x):,}")
        st.table(df_show)
        total_amt = sum(e["amount"] for e in st.session_state.qoo10_entries)
        total_qty = sum(e["qty"] for e in st.session_state.qoo10_entries)
        st.caption(f"합계: {len(st.session_state.qoo10_entries)}건 / 수량 {int(total_qty):,} / 금액 {int(total_amt):,} JPY")
        if st.button("목록 전체 지우기"):
            st.session_state.qoo10_entries = []
            # 지운 결과(0건·빈 상태 문구)를 보여줘야 처리됐다는 확인이 됩니다.
            st.session_state.qoo10_panel_open = True
            st.rerun()
    else:
        st.caption("아직 추가된 큐텐재팬 건이 없습니다.")

# ══════════════════════════════════════════════════════════════════
# STEP 3 — 신고기간 확인
# ══════════════════════════════════════════════════════════════════
step_head(3, "신고기간 확인")

_period_mode = st.radio(
    "신고기간 지정 방식",
    ["전체 (기간 제한 없음)", "반기·분기", "직접 입력"],
    horizontal=True,
    label_visibility="collapsed",
)

report_period_start = ""
report_period_end = ""
_this_year = date.today().year

if _period_mode == "반기·분기":
    pc1, pc2 = st.columns([1, 2])
    _year = pc1.number_input("연도", min_value=2000, max_value=2100,
                             value=_this_year, step=1, format="%d")
    _presets = period_presets(int(_year))
    _label = pc2.selectbox("기간", list(_presets))
    report_period_start, report_period_end = _presets[_label]
elif _period_mode == "직접 입력":
    pc1, pc2 = st.columns(2)
    _start = pc1.date_input("시작일", value=date(_this_year, 1, 1), format="YYYY-MM-DD")
    _end = pc2.date_input("종료일", value=date(_this_year, 6, 30), format="YYYY-MM-DD")
    if _start > _end:
        st.error("시작일이 종료일보다 늦습니다.")
    report_period_start, report_period_end = _start.isoformat(), _end.isoformat()

if report_period_start and report_period_end:
    note(f"신고기간 <b>{html.escape(report_period_start)} ~ {html.escape(report_period_end)}</b> — "
         "선적일자(기적일)가 이 범위 안인 거래만 반영합니다. 시작일과 종료일은 모두 포함하며, "
         "범위 밖 거래는 사유와 함께 처리 로그에 남깁니다.")
else:
    note("업로드한 자료를 기간 제한 없이 모두 반영합니다. 신고기간을 지정하면 선적일자 기준으로 "
         "해당분만 반영하고 이전기간·미도래·날짜없음 건을 따로 보여줍니다.")

# ══════════════════════════════════════════════════════════════════
# STEP 4 — 생성 문서 선택 및 환율 안내
# ══════════════════════════════════════════════════════════════════
step_head(4, "만들 문서 고르기")
cc1, cc2, cc3 = st.columns(3)
make_sales = cc1.checkbox("매출집계", value=True)
make_zero = cc2.checkbox("영세율첨부서류제출명세서", value=True)
make_export = cc3.checkbox("수출실적명세서", value=False)

zero_doc_mode = "전체"
if make_zero:
    zero_doc_mode = st.radio(
        "영세율첨부서류제출명세서 생성 범위",
        ["전체", "월별"],
        horizontal=True,
        help="전체를 선택하면 전체 통합 파일 1개만, 월별을 선택하면 월별 파일만 생성합니다.",
    )

note("환율은 서울외국환중개 매매기준율에서 자동으로 가져옵니다. 이미 받아 둔 구간은 캐시를 쓰고 모자란 구간만 새로 조회합니다.")

# ══════════════════════════════════════════════════════════════════
# STEP 5 — 처리 시작
# ══════════════════════════════════════════════════════════════════
step_head(5, "생성")
has_process_input = bool(uploaded_files) or bool(st.session_state.qoo10_entries)
process_btn = st.button(
    "엑셀 파일 생성",
    type="primary",
    use_container_width=True,
    disabled=not has_process_input,
)
if not has_process_input:
    st.caption("PDF·Excel·CSV를 업로드하거나 큐텐재팬 정보를 입력하면 생성 버튼이 활성화됩니다.")

progress_bar = st.empty()
status_text = st.empty()
log_area = st.empty()


def _needed_currencies(shopee_results, lazada_result, qoo10_result,
                       joom_results=None, shopify_results=None):
    used = set()
    for sd in shopee_results or []:
        if sd.get("currency"):
            used.add(sd["currency"])
    if lazada_result:
        for it in lazada_result.get("items", []):
            if it.get("currency"):
                used.add(it["currency"])
    # Joom·쇼피파이는 건별 기준일의 일별 매매기준율을 사용합니다.
    for res in list(joom_results or []) + list(shopify_results or []):
        for it in res.get("items", []):
            if it.get("currency"):
                used.add(it["currency"])
    # 큐텐재팬은 일별/기간평균 환율을 사용하지 않고 반기말(6월/12월)의
    # 서울외국환중개 공식 월평균 매매기준율만 사용합니다.
    # 따라서 큐텐 때문에 JPY 일별 환율을 별도로 수집하지 않습니다.
    return sorted(used)



def _qoo10_reporting_month(entry=None, result=None):
    """큐텐 거래기간 기준 반기말 월(YYYY-06 또는 YYYY-12)을 반환합니다."""
    entry = entry or {}
    result = result or {}

    period_end = entry.get("period_end") or result.get("period_end") or ""
    digits = re.sub(r"\D", "", str(period_end))[:8]
    if len(digits) >= 6:
        year = digits[:4]
        month = int(digits[4:6])
        return f"{year}-06" if month <= 6 else f"{year}-12"

    base = (
        entry.get("period_start") or result.get("period_start")
        or entry.get("write_date") or result.get("write_date") or ""
    )
    digits = re.sub(r"\D", "", str(base))[:8]
    if len(digits) >= 6:
        year = digits[:4]
        month = int(digits[4:6])
        return f"{year}-06" if month <= 6 else f"{year}-12"
    return ""


def _monthly_rate_requests(ebay_results, qoo10_result):
    """통화별로 공식 월평균 환율이 필요한 월 목록을 만듭니다."""
    requests = {}

    # 이베이: PDF의 실제 발행월별 월평균 환율
    for er in ebay_results or []:
        for it in er.get("items", []):
            currency = str(it.get("currency", "")).strip().upper()
            month_value = str(it.get("month", "")).strip()
            if currency and re.fullmatch(r"20\d{2}-\d{2}", month_value):
                requests.setdefault(currency, set()).add(month_value)

    # 큐텐재팬: 거래기간이 속한 반기의 말월(6월 또는 12월) 월평균 환율
    if qoo10_result:
        entries = qoo10_result.get("entries") or [{}]
        for entry in entries:
            month_value = _qoo10_reporting_month(entry, qoo10_result)
            if month_value:
                requests.setdefault("JPY", set()).add(month_value)

    return {currency: sorted(months) for currency, months in requests.items() if months}


def _filter_monthly_rate_data(rate_data, requested_months):
    """조회 구간 중 실제 필요한 월만 환율 시트에 남깁니다."""
    wanted = set(requested_months or [])
    filtered = [
        row for row in (rate_data.get("monthly", []) or [])
        if str(row.get("year_month", "")) in wanted
    ]
    values = [float(row.get("rate", 0) or 0) for row in filtered if float(row.get("rate", 0) or 0) > 0]
    result = dict(rate_data or {})
    result["monthly"] = filtered
    result["period"] = (
        f"{requested_months[0]} ~ {requested_months[-1]}" if requested_months else ""
    )
    result["average"] = round(sum(values) / len(values), 2) if values else 0.0
    result["monthly_average"] = result["average"]
    result["min"] = min(values) if values else 0.0
    result["max"] = max(values) if values else 0.0
    result["range"] = round(max(values) - min(values), 2) if values else 0.0
    return result

def _daily_rate_period_bounds(shopee_results, lazada_result, qoo10_result,
                              joom_results=None, shopify_results=None):
    """실제 신고기간의 시작/종료일을 반환합니다. 작성일은 환율시트 기간에 포함하지 않습니다."""
    starts = []
    ends = []

    def _add(start_value, end_value):
        sdt = pd.to_datetime(str(start_value or "").replace(".", "-"), errors="coerce")
        edt = pd.to_datetime(str(end_value or "").replace(".", "-"), errors="coerce")
        if not pd.isna(sdt):
            starts.append(sdt.normalize())
        if not pd.isna(edt):
            ends.append(edt.normalize())

    for sd in shopee_results or []:
        _add(sd.get("period_start"), sd.get("period_end"))
        if not sd.get("period_start") or not sd.get("period_end"):
            tx_dates = [pd.to_datetime(str(tx.get("date", "")).replace(".", "-"), errors="coerce") for tx in sd.get("transactions", [])]
            tx_dates = [d.normalize() for d in tx_dates if not pd.isna(d)]
            if tx_dates:
                starts.append(min(tx_dates)); ends.append(max(tx_dates))

    if lazada_result:
        _add(lazada_result.get("period_start"), lazada_result.get("period_end"))

    if qoo10_result:
        _add(qoo10_result.get("period_start"), qoo10_result.get("period_end"))
        for entry in qoo10_result.get("entries", []):
            _add(entry.get("period_start"), entry.get("period_end"))

    # Joom·쇼피파이는 문서 기간이 아니라 실제 건별 기준일 범위를 사용합니다.
    for res in list(joom_results or []) + list(shopify_results or []):
        item_dates = [str(it.get("date", "")) for it in res.get("items", []) if it.get("date")]
        if item_dates:
            _add(min(item_dates), max(item_dates))
        else:
            _add(res.get("period_start"), res.get("period_end"))

    if not starts and not ends:
        return None, None
    display_start = min(starts or ends)
    display_end = max(ends or starts)
    return display_start, display_end


def _build_qoo10_result():
    entries = list(st.session_state.qoo10_entries)
    if not entries:
        return None
    submitter = next((e.get("_submitter") for e in entries if (e.get("_submitter") or {}).get("name")), {})
    return {
        "submitter": submitter,
        "type": "qoo10",
        "carrier": "국제로지스틱",
        "destination": "JP",
        "currency": "JPY",
        "period_start": min((e.get("period_start") for e in entries if e.get("period_start")), default=""),
        "period_end": max((e.get("period_end") for e in entries if e.get("period_end")), default=""),
        "write_date": max((e.get("write_date") for e in entries if e.get("write_date")), default=""),
        "qty": sum(int(e.get("qty", 0) or 0) for e in entries),
        "amount": sum(float(e.get("amount", 0) or 0) for e in entries),
        "tracking_no": entries[0].get("tracking_no", "") if entries else "",
        "entries": entries,
    }

if process_btn:
    if not (make_sales or make_zero or make_export):
        st.error("생성할 문서를 하나 이상 선택해 주세요.")
        st.stop()

    logs = []
    def log(msg):
        logs.append(str(msg))
        log_area.markdown('<div class="log-box">' + "\n".join(logs[-120:]) + '</div>', unsafe_allow_html=True)

    def log_refund_warnings(result):
        """환불 감지 장치가 남긴 경고를 처리 로그에 표시합니다 (전 플랫폼 공통)."""
        for w in (result or {}).get("refund_warnings") or []:
            log(f"[WARN] {w}")

    st.session_state.result_files = []
    progress_bar.progress(3, text="처리 준비 중...")
    status_text.info("파일 분석, 환율 수집, 엑셀 생성 중입니다...")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        input_paths = []
        seen_digests = {}
        for i, uf in enumerate(uploaded_files or []):
            payload = bytes(uf.getbuffer())
            # ① 같은 바이트의 파일은 한 번만 반영 (이름이 달라도 SHA-256으로 판정)
            digest = hashlib.sha256(payload).hexdigest()
            if digest in seen_digests:
                log(f"[중복] 같은 내용의 파일이라 한 번만 반영: {uf.name} (= {seen_digests[digest]})")
                continue
            seen_digests[digest] = uf.name
            # 같은 파일명이 다른 내용으로 올라와도 덮어쓰지 않도록 업로드별 하위 폴더에 저장
            sub = tmpdir / f"upload_{i:02d}"
            sub.mkdir(exist_ok=True)
            p = sub / uf.name
            p.write_bytes(payload)
            input_paths.append(p)

        try:
            # PDF / 라자다 주문내역 Excel 파싱
            t_pdf = time.perf_counter()
            progress_bar.progress(15, text="📄 파일 분석 중...")
            log("📄 파일 분석 중...")
            shopee_results = []
            lazada_results = []
            ebay_results = []
            joom_results = []
            shopify_results = []
            for p in input_paths:
                suffix = p.suffix.lower()
                selected_type = uploaded_type_choices.get(p.name, "")

                if suffix in {".xlsx", ".xlsm", ".csv"}:
                    if selected_type == "shopify":
                        result = parse_shopify_orders(p)
                        shopify_results.append(result)
                        totals = result.get("total_by_currency", {})
                        total_text = ", ".join(f"{amount:,.2f} {cur}" for cur, amount in totals.items())
                        reasons = result.get("skipped_by_reason") or {}
                        reason_text = ", ".join(
                            f"{k} {v['count']}건" for k, v in sorted(reasons.items())
                        )
                        log(f"[OK] 쇼피파이 {result.get('store','')}: {p.name} / "
                            f"{result.get('row_count', 0):,}건 / {total_text}"
                            + (f" / 미반영: {reason_text}" if reason_text else ""))
                        log_refund_warnings(result)
                        continue
                    if selected_type != "lazada_excel":
                        log(f"[WARN] 지원하지 않는 Excel/CSV 형식: {p.name}")
                        continue
                    result = parse_lazada_order_excel(p)
                    lazada_results.append(result)
                    totals = result.get("total_amount_by_currency", {})
                    total_text = ", ".join(f"{amount:,.2f} {cur}" for cur, amount in totals.items())
                    reasons = result.get("skipped_by_reason") or {}
                    reason_text = ", ".join(f"{k} {v['count']}건" for k, v in sorted(reasons.items()))
                    log(f"[OK] 라자다 주문엑셀: {p.name} / {len(result.get('items', [])):,}건 / {total_text}"
                        + (f" / 미반영: {reason_text}" if reason_text else ""))
                    log_refund_warnings(result)
                    continue

                detected_from_file = detect_pdf_type(str(p))
                forced_type = selected_type if detected_from_file == "unknown" else None
                detected_type = forced_type or detected_from_file
                # 큐텐 PDF는 업로드 단계에서 파싱되어 STEP 2 목록에 자동 반영됩니다.
                if detected_type == "qoo10":
                    matched = [e for e in st.session_state.qoo10_entries if e.get("_source_file") == p.name]
                    if matched:
                        e = matched[0]
                        log(f"[OK] 큐텐재팬: {p.name} / {int(e.get('qty', 0)):,}건 / {int(e.get('amount', 0)):,} JPY")
                        for w in e.get("_warnings") or []:
                            log(f"[WARN] {w}")
                    else:
                        log(f"[WARN] 큐텐재팬 PDF 자동입력 실패: {p.name} / STEP 2에서 직접 입력")
                    continue
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    result = parse_pdf(str(p), forced_type=forced_type)
                if not result:
                    log(f"[WARN] 파싱 실패 또는 미확인: {p.name}")
                    continue
                if result.get("type") == "shopee":
                    shopee_results.append(result)
                    log(f"[OK] 쇼피 {result.get('currency','?')}: {p.name} / {result.get('total_qty',0):,}건")
                elif result.get("type") == "lazada":
                    result.setdefault("source_files", [p.name])
                    result.setdefault("source_kind", "receipt_pdf")
                    lazada_results.append(result)
                    log(f"[OK] 라자다 PDF: {p.name} / {len(result.get('items', [])):,}건")
                elif result.get("type") == "ebay":
                    ebay_results.append(result)
                    currencies = sorted({it.get("currency") for it in result.get("items", []) if it.get("currency")})
                    log(f"[OK] 이베이/린코스: {p.name} / {len(result.get('items', [])):,}건 / {', '.join(currencies)}")
                elif result.get("type") == "joom":
                    joom_results.append(result)
                    totals = result.get("total_by_currency", {})
                    total_text = ", ".join(f"{amount:,.2f} {cur}" for cur, amount in totals.items())
                    log(f"[OK] Joom: {p.name} / {len(result.get('items', [])):,}건 / {total_text}")
                    if result.get("total_mismatch"):
                        log(f"[WARN] Joom 합계 불일치: PDF 표기 {result.get('declared_total')} / 건별 {totals}")
                log_refund_warnings(result)

            lazada_result = merge_lazada_results(lazada_results)
            shopify_results = merge_shopify_results(shopify_results)
            qoo10_result = _build_qoo10_result()
            if qoo10_result:
                log(f"[OK] 큐텐 STEP 2: {len(qoo10_result.get('entries', [])):,}건 / {int(qoo10_result.get('amount',0)):,} JPY")

            if not (shopee_results or lazada_result or qoo10_result or ebay_results
                    or joom_results or shopify_results):
                raise RuntimeError("처리할 데이터가 없습니다. PDF/Excel/CSV 또는 큐텐 수동 입력을 확인해 주세요.")
            log(f"✅ 파일 분석 완료 ({time.perf_counter() - t_pdf:.1f}초)")

            # ② 거래 중복·충돌 검사 — 확정 중복(키·내용 모두 동일)만 자동 제외하고,
            # 같은 거래키에 내용이 다른 자료(변경 충돌)는 자동 선택하지 않고 중단합니다.
            try:
                guard = dedup_transactions(
                    shopee_results=shopee_results, lazada_result=lazada_result,
                    qoo10_result=qoo10_result, ebay_results=ebay_results,
                    joom_results=joom_results, shopify_results=shopify_results,
                )
            except DuplicateConflictError as exc:
                for line in str(exc).splitlines():
                    log(f"[STOP] {line}")
                raise
            shopee_results = guard["shopee_results"]
            lazada_result = guard["lazada_result"]
            qoo10_result = guard["qoo10_result"]
            ebay_results = guard["ebay_results"]
            joom_results = guard["joom_results"]
            shopify_results = guard["shopify_results"]
            for line in format_dedup_lines(guard["report"]):
                log(line)

            # ③ 사업자 혼합 차단 — 읽힌 사업자번호가 2개 이상이면 중단합니다.
            try:
                for w in check_submitter_mix(
                    shopee_results=shopee_results, lazada_result=lazada_result,
                    qoo10_result=qoo10_result, ebay_results=ebay_results,
                    joom_results=joom_results, shopify_results=shopify_results,
                ):
                    log(f"[WARN] {w}")
            except MixedSubmitterError as exc:
                for line in str(exc).splitlines():
                    log(f"[STOP] {line}")
                raise

            # 신고기간 분류 — 매출집계·신고서류·환율 범위가 모두 같은 거래집합을 쓰도록
            # 소비자로 갈라지기 전 이 지점에서 한 번만 거릅니다.
            try:
                scoped = apply_reporting_period(
                    shopee_results=shopee_results, lazada_result=lazada_result,
                    qoo10_result=qoo10_result, ebay_results=ebay_results,
                    joom_results=joom_results, shopify_results=shopify_results,
                    start=report_period_start, end=report_period_end,
                )
            except PeriodStraddleError as exc:
                for line in str(exc).splitlines():
                    log(f"[STOP] {line}")
                raise
            shopee_results = scoped["shopee_results"]
            lazada_result = scoped["lazada_result"]
            qoo10_result = scoped["qoo10_result"]
            ebay_results = scoped["ebay_results"]
            joom_results = scoped["joom_results"]
            shopify_results = scoped["shopify_results"]
            for line in format_exclusion_lines(scoped["report"]):
                log(line)
            if not (shopee_results or lazada_result or qoo10_result or ebay_results
                    or joom_results or shopify_results):
                raise RuntimeError(
                    f"신고기간({report_period_start} ~ {report_period_end})에 해당하는 거래가 "
                    "없습니다. 기간을 다시 확인하거나 STEP 3에서 '전체'를 선택해 주세요."
                )

            # 환율 수집
            daily_needed = _needed_currencies(shopee_results, lazada_result, qoo10_result,
                                              joom_results=joom_results, shopify_results=shopify_results)
            monthly_requests = _monthly_rate_requests(ebay_results, qoo10_result)
            display_start, display_end = _daily_rate_period_bounds(
                shopee_results, lazada_result, qoo10_result,
                joom_results=joom_results, shopify_results=shopify_results,
            )
            if display_start is None or display_end is None:
                today = pd.Timestamp.today().normalize()
                display_start = today
                display_end = today
            # 1월 1일 등 휴일의 직전 영업일 환율을 확보하기 위해 7일 앞에서부터 수집합니다.
            rate_start = display_start - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
            rate_end = display_end

            t_rate = time.perf_counter()
            progress_bar.progress(45, text="💱 환율 확인 중...")
            rates = {}
            if daily_needed:
                rates = fetch_all_currencies_for_period(
                    rate_start, rate_end, daily_needed, logger=log,
                    display_start=display_start, display_end=display_end,
                )
            if monthly_requests:
                requested_currencies = sorted(monthly_requests)
                log(f"💱 월평균 환율 확인 중... ({', '.join(requested_currencies)})")
                monthly_rates = {}
                silent_logger = lambda _msg: None
                for currency in requested_currencies:
                    requested_months = monthly_requests[currency]
                    fetched = fetch_monthly_avg_currencies_for_period(
                        requested_months[0], requested_months[-1], [currency], logger=silent_logger
                    )
                    monthly_rates[currency] = _filter_monthly_rate_data(
                        fetched[currency], requested_months
                    )
                rates = merge_monthly_rates(rates, monthly_rates)
                log("✅ 월평균 환율 확인 완료")
            log(f"✅ 환율 확인 완료 ({time.perf_counter() - t_rate:.1f}초)")

            # 출력 라벨/파일명
            year = rate_end.year
            month = rate_end.month
            disp_label, fname_label = period_labels(
                shopee_results, lazada_result, qoo10_result, ebay_results=ebay_results,
                joom_results=joom_results, shopify_results=shopify_results,
                fallback=f"{year}년 {month:02d}월",
            )
            fsafe = safe_filename(fname_label or f"{year}{month:02d}")
            company = company_name_from_results(
                shopee_results, lazada_result, qoo10_result, ebay_results=ebay_results,
                joom_results=joom_results, shopify_results=shopify_results,
            )

            created = []
            t_excel = time.perf_counter()
            progress_bar.progress(75, text="📊 엑셀 생성 중...")
            log("📊 선택한 문서 생성 중...")

            if make_sales:
                sales_path = tmpdir / f"매출집계_{fsafe}.xlsx"
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    generate_excel(
                        shopee_results=shopee_results,
                        lazada_result=lazada_result,
                        qoo10_result=qoo10_result,
                        rates=rates,
                        output_path=str(sales_path),
                        ebay_results=ebay_results,
                        joom_results=joom_results,
                        shopify_results=shopify_results,
                        year=year,
                        month=month,
                    )
                created.append(sales_path)
                log(f"[OK] 매출집계 생성: {sales_path.name}")

            if make_zero or make_export:
                rows = build_declaration_rows(
                    shopee_results, lazada_result, qoo10_result, rates,
                    ebay_results=ebay_results,
                    joom_results=joom_results, shopify_results=shopify_results,
                )
                if make_zero:
                    zero_mode_arg = "all" if zero_doc_mode == "전체" else "monthly"
                    zero_files = create_zero_rate_attachments(
                        rows,
                        tmpdir,
                        company,
                        base_dir=BASE_DIR,
                        mode=zero_mode_arg,
                    )
                    created.extend(zero_files)
                    log(f"[OK] 영세율첨부서류제출명세서 생성({zero_doc_mode}): {len(zero_files)}개")
                if make_export:
                    export_file = create_export_performance(rows, tmpdir, company, base_dir=BASE_DIR)
                    created.append(export_file)
                    log(f"[OK] 수출실적명세서 생성: {export_file.name}")

            result_files = []
            for p in created:
                result_files.append({
                    "name": p.name,
                    "bytes": p.read_bytes(),
                    "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                })
            st.session_state.result_files = result_files
            log(f"✅ 문서 생성 완료 ({time.perf_counter() - t_excel:.1f}초)")
            progress_bar.progress(100, text="✅ 완료")
            status_text.success(f"✅ 엑셀 생성 완료! — {disp_label}")
            log("✅ 전체 처리 완료")
        except Exception as e:
            progress_bar.progress(100, text="오류 발생")
            status_text.error(f"❌ 오류: {e}")
            st.exception(e)

if st.session_state.result_files:
    step_head(6, "내려받기")
    for i, f in enumerate(st.session_state.result_files):
        st.download_button(
            f["name"],
            data=f["bytes"],
            file_name=f["name"],
            mime=f["mime"],
            key=f"download_{i}_{f['name']}",
            use_container_width=True,
        )

st.markdown('<div style="border-top:1px solid var(--line); margin:2rem 0 1rem"></div>',
            unsafe_allow_html=True)
with st.expander("파일명 규칙"):
    st.markdown(
        """
| 파일명 패턴 | 플랫폼 |
|---|---|
| `업체명_MY_*.pdf` | 쇼피 말레이시아 |
| `업체명_PH_*.pdf` | 쇼피 필리핀 |
| `업체명_SG_*.pdf` | 쇼피 싱가폴 |
| `업체명_TH_*.pdf` | 쇼피 태국 |
| `업체명_TW_*.pdf` | 쇼피 대만 |
| `업체명_VN_*.pdf` | 쇼피 베트남 |
| `라자다_*.pdf` | 라자다 소포수령증 |
| `라자다_*.xlsx` | 라자다 주문내역 (`paidPrice`/`deliveredDate`) |
| `큐텐재팬_*.pdf` | 큐텐재팬 — PDF 자동인식 후 STEP 2에 반영 |
| `이베이_*.pdf` / `린코스_*.pdf` | 린코스(주) 소포수령증 — 발행월별 다통화 |
| `*.pdf` (본문에 `H3 NETWORKS`) | Joom — 에이치3네트웍스 상품 수령 및 운송 확인증 |
| `<스토어> orders *.csv` | 쇼피파이 주문내역 (`Fulfilled at`/`Total`) |

참고
- 쇼피는 업체명과 무관하게 `_MY_`, `_PH_`, `_SG_`, `_TH_`, `_TW_`, `_VN_`, `_BR_`, `_MX_` 국가코드 패턴도 함께 인식합니다.
- Excel/CSV는 파일명이 아니라 **열 구성**으로 판별하므로 파일명을 바꿔도 됩니다.
- 쇼피파이는 `transaction`, `payout` 파일이 아니라 **orders** 파일을 올려야 합니다.
- 적용환율: 쇼피·라자다(주문내역)·Joom·쇼피파이는 **건별 기준일 일별환율**,
  이베이(린코스)·큐텐재팬은 **공식 월평균 매매기준율** 입니다.
"""
    )
