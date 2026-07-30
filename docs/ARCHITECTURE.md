# 아키텍처 — 데이터 흐름 · 결과 스키마 · 신규 플랫폼 추가 절차

---

## 1. 모듈 책임

| 모듈 | 책임 | 외부 의존 |
|---|---|---|
| `app.py` | Streamlit UI, 업로드 파일 판별, 파싱 오케스트레이션, 환율 요청 범위 계산 | streamlit |
| `process.py` | 동일 로직의 CLI 배치판 (`input/` → `output/`) | yaml |
| `modules/pdf_parser.py` | PDF 종류 판별 + 쇼피/라자다/큐텐/린코스/Joom 파서 | pdfplumber, (pytesseract) |
| `modules/lazada_order_parser.py` | 라자다 주문내역 xlsx 파서 | openpyxl, pandas |
| `modules/shopify_parser.py` | 쇼피파이 orders CSV/xlsx 파서 | openpyxl, pandas |
| `modules/exchange_rate.py` | SMBS 일별/월평균 수집·캐시·조회·단위 정규화 | requests, bs4, (selenium) |
| `modules/reporting_period.py` | 신고기간 비파괴 분류 (해당분/이전기간/미도래/날짜없음/분할불가 차단) | - |
| `modules/dedup_guard.py` | 중복·변경충돌·사업자 혼합 검사 (기준: docs/POLICY.md) | - |
| `modules/excel_writer.py` | 매출집계 워크북 생성 | openpyxl |
| `modules/extra_docs.py` | 영세율첨부서류제출명세서 / 수출실적명세서 | openpyxl |

파싱 결과는 소비자(매출집계/신고서류)로 갈라지기 **전에** `dedup_guard` →
`check_submitter_mix` → `reporting_period` 순으로 한 번만 거른다
(`app.py`·`process.py` 동일). 필터를 소비자별로 두 번 구현하면 산출물 합계가
어긋나므로 이 순서를 깨지 말 것.

---

## 2. 파서 결과 dict 스키마

모든 파서는 `type` 키로 자신을 식별합니다.

### 쇼피 — `type: 'shopee'`
```python
{
  'type': 'shopee', 'submitter': {...}, 'carrier': str,
  'country': 'MY', 'currency': 'MYR',
  'period_start': 'YYYY-MM-DD', 'period_end': ..., 'write_date': ...,
  'total_qty': int, 'total_amount': float,
  'transactions': [{'carrier','date','tracking_no','country','qty','amount'}],
}
```

### 라자다 — `type: 'lazada'`
```python
{
  'type': 'lazada', 'source_kind': 'receipt_pdf'|'order_excel'|'mixed',
  'source_files': [str], 'source_workbooks': [{'filename','sheet_name','content': bytes}],
  'carrier','period_start','period_end','write_date','submitter',
  'items': [{'service','carrier','origin','destination','tracking_no',
             'qty','amount','currency','date','delivered_date','source_kind', ...}],
}
```
> 여러 결과는 `merge_lazada_results()` 로 하나의 dict 로 합칩니다.

### 큐텐 — `type: 'qoo10'`
```python
{
  'type': 'qoo10', 'submitter': {...}, 'carrier': '국제로지스틱',
  'destination': 'JP', 'currency': 'JPY',
  'period_start','period_end','write_date',
  'qty': int, 'amount': float, 'tracking_no': str,
  'entries': [{'period_start','period_end','tracking_no','qty','amount','write_date', ...}],
}
```
> `entries` 는 STEP 2 화면의 입력 목록. `generate_excel()` 이 건별 `rate`/`krw`/`rate_month` 를 채워 넣습니다.

### 린코스(이베이) — `type: 'ebay'`
```python
{
  'type': 'ebay', 'platform': '이베이', 'carrier': '린코스(주)', 'submitter': {...},
  'period_start','period_end','write_date',
  'items': [{'carrier','service','country','destination','tracking_no',
             'month': 'YYYY-MM', 'date': 월말, 'period_start','period_end',
             'currency','qty','amount', 'rate_basis': 'monthly_average'}],
  'summary_items': [ ... ],   # 표시 전용
}
```

### Joom — `type: 'joom'`
```python
{
  'type': 'joom', 'platform': 'Joom', 'carrier': '에이치3네트웍스', 'submitter': {...},
  'period_start','period_end','write_date',
  'items': [{'service','date','order_id','destination','tracking_no',
             'currency': 'USD', 'qty': 1, 'amount': float, 'item_name': str,
             'rate_basis': 'daily'}],
  'declared_total': {'USD': 254.83},   # PDF 합계 행
  'total_by_currency': {'USD': 254.83},
  'total_mismatch': bool,
}
```

### 쇼피파이 — `type: 'shopify'`
```python
{
  'type': 'shopify', 'platform': '쇼피파이', 'store': 'FRESORY',
  'source_file': 'FRESORY orders 26년.csv',
  'headers': [원본 열 이름...],          # 원본 시트 재현용
  'rows': [[원본 셀 값...], ...],        # 원본 시트 재현용 (전 행)
  'row_flags': [{'index': int, 'date': 'YYYY-MM-DD'|'', 'counted': bool}],
  'period_start','period_end','write_date',
  'items': [{'order_name','date','currency','qty','amount',
             'financial_status','fulfilled_at','row_index','rate_basis': 'daily'}],
  'total_by_currency': {'USD': 517.76},
  'skipped_unfulfilled': int, 'skipped_blank': int,
}
```

---

## 3. 환율 적용 규칙 매트릭스

| 플랫폼 | 기준일 | 조회 함수 |
|---|---|---|
| 쇼피 | 거래 `date` | `get_rate_for_date()` |
| 라자다 (PDF) | 거래기간 평균 | `avg_rate_for_period()` |
| 라자다 (Excel) | `deliveredDate` | `get_rate_for_date()` |
| Joom | 발송날짜 | `get_rate_for_date()` |
| 쇼피파이 | `Fulfilled at` | `get_rate_for_date()` |
| 린코스(이베이) | 발행월 | `monthly_avg_rate_for_month()` |
| 큐텐 | 반기말 월 | `monthly_avg_rate_for_month()` |

* 일별: 요청 구간보다 `RATE_LOOKBACK_DAYS=7` 일 앞에서부터 수집 → 휴일은 **직전 영업일** 값으로 ffill
* 월평균: SMBS 월평균 페이지의 **공식 값만** 사용. 없으면 `RuntimeError` 로 중단 (자체 평균 금지)
* 단위: `normalize_smbs_rate()` 가 JPY·IDR·VND 의 100통화 고시값을 1통화로 환산.
  `smbs_source_rate()` 는 반대로 환율 시트 표시용 100통화 값으로 되돌림.

---

## 4. 워크북 시트 구성 (`generate_excel`)

```
총집계            플랫폼 × 통화 외화/원화 합계
월별집계          월 × 구분 × 통화
<통화>            수출신고 프로그램 업로드용 (수출신고번호/기타영세율건수/선적일자/통화/환율/외화/원화)
쇼피(MYR)…        플랫폼 원본 시트
라자다(주문내역)
이베이 / Joom / 쇼피파이(스토어)
큐텐(소포수령증)
환율(<통화>)      적용환율 근거
```

마지막에 `_prune_workbook_sheets()` 가 `keep_sheets` 에 없는 시트를 삭제하므로,
**시트를 새로 만들면 반드시 `keep_sheets` 에도 추가**해야 합니다.

---

## 5. 신규 플랫폼 추가 체크리스트

새 배송업체/마켓 자료를 붙일 때 손대야 하는 지점입니다.

### 5-1. 파싱
- [ ] `modules/<platform>_parser.py` 또는 `pdf_parser.py` 에 파서 추가
- [ ] `pdf_parser.detect_pdf_type()` / `_detect_pdf_type_from_text()` 에 판별 규칙 추가
- [ ] `pdf_parser.parse_pdf()` 의 분기와 `aliases` 에 추가

### 5-2. UI / 오케스트레이션 (`app.py`)
- [ ] `UNKNOWN_PDF_TYPE_OPTIONS` / `UNKNOWN_PDF_TYPE_TO_CODE`
- [ ] `_detect_uploaded_file_type()` (Excel/CSV면 헤더 판별 추가)
- [ ] 업로드 목록의 아이콘/라벨 dict
- [ ] 처리 루프에서 결과 수집 리스트
- [ ] `_needed_currencies()` (일별 환율이 필요한 통화)
- [ ] `_monthly_rate_requests()` (월평균이 필요한 통화·월)
- [ ] `_daily_rate_period_bounds()` (환율 조회 구간)
- [ ] `generate_excel(...)` / `build_declaration_rows(...)` 호출 인자
- [ ] `modules/reporting_period.py` — 새 플랫폼의 기간 필터 함수(건별 날짜형이면
      `_filter_pointwise` 재사용, 집계형이면 `classify_span` 기반 차단) + `apply_reporting_period` 연결
- [ ] `modules/dedup_guard.py` — 강한 거래키 정의(실측으로 유일성 확인 후) + `dedup_transactions` 연결.
      키 유일성이 불확실하면 자동 제외 대신 경고만 (docs/POLICY.md 4단계 기준)

### 5-3. 엑셀 출력 (`modules/excel_writer.py`)
- [ ] `period_labels()`
- [ ] `_infer_used_sources_and_currencies()`
- [ ] `write_summary_sheet()` — 총집계 블록
- [ ] `write_monthly_summary_sheet()` — 월별집계 행
- [ ] `write_currency_template_sheet()` — 통화별 시트 요약행 + 데이터행
- [ ] `write_<platform>_sheet()` 신설
- [ ] `generate_excel()` — 시트 생성 + `keep_sheets` 등록
- [ ] `PREFERRED_CURRENCY_ORDER` / `COUNTRY_NAMES` 에 새 통화 추가

### 5-4. 신고서류 (`modules/extra_docs.py`)
- [ ] `build_declaration_rows()` 에 행 생성
- [ ] `_company_name()` 에 제출자 추출 순서 추가

### 5-5. CLI (`process.py`)
- [ ] `collect_inputs()` 확장자, 파싱 분기, `generate_excel()` 인자

### 5-6. 문서
- [ ] `docs/PLATFORMS.md` 에 포맷·판별·환율 규칙 기록
- [ ] `README.md` 지원 플랫폼 표

---

## 6. 알려진 제약

* `write_currency_template_sheet()` 는 **1~5행이 플랫폼별 요약(쇼피·라자다·이베이·Joom·쇼피파이),
  6행이 헤더, 7행부터 데이터** 입니다. 플랫폼을 추가하면 `summary_rows` 에 항목을 넣으면
  헤더/데이터 행이 자동으로 밀립니다.
* `JPY` 시트는 큐텐 전용이라 `generate_excel()` 안에서 따로 작성하며 헤더가 **4행** 입니다.
  다른 통화 시트(6행 헤더)와 위치가 다르니 데이터를 복사할 때 주의하세요.
* 월평균 매매기준율을 못 얻으면 `RuntimeError` 로 중단되며, **일별 환율 평균으로 대체하지
  않습니다.** 재시도하거나 `data/monthly_exchange_rate_cache.csv` 에 공식값을
  `source=SMBS_MON_AVG_OFFICIAL` 로 넣어야 합니다.
* **큐텐만 환율 예외**: 다른 플랫폼과 달리 반기말(6월/12월)에 환율을 모아 적용합니다.
  월 단위 큐텐 PDF(예: 10월분)도 선적일자는 그 달 말일이지만 **환율은 반기말 월평균**입니다.
  이 규칙은 세 곳이 같은 답을 내야 합니다 — `app.py::_qoo10_reporting_month`(수집),
  `excel_writer.py::_qoo10_reporting_month`(조회), `extra_docs.build_declaration_rows`(신고행).
  한쪽만 바꾸면 '수집은 12월, 조회는 10월'로 어긋나 `RuntimeError` 가 납니다.
* 큐텐 이미지 PDF OCR은 `pdf2image` + `pytesseract` + poppler/tesseract 바이너리가 있어야 동작합니다.
  없으면 STEP 2 수동 입력으로 처리합니다.
* SMBS 사이트가 requests 로 막히면 Selenium(Chromium) 폴백을 씁니다. 컨테이너에 브라우저가 없으면 실패합니다.
