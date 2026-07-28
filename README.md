# 📦 소포수령증 자동화 (sopo)

해외 오픈마켓·배송대행사가 발급한 **해외배송 소포수령증(PDF)** 과 **주문내역(Excel/CSV)** 을 읽어
부가가치세 신고용 **매출집계 / 영세율첨부서류제출명세서 / 수출실적명세서** 를 자동 생성하는 Streamlit 앱입니다.

환율은 **서울외국환중개(SMBS)** 에서 직접 수집하며, 수집한 값은 `data/` 아래 CSV 캐시에 저장해
같은 구간을 반복 조회하지 않습니다.

---

## 1. 빠른 시작

```bash
python -m pip install -r requirements.txt
```

웹 앱 실행:

```bash
python -m streamlit run app.py
```

CLI 배치 실행 (`input/` 폴더의 파일을 일괄 처리):

```bash
python process.py --year 2026 --month 03
```

Docker:

```bash
docker build -t sopo . && docker run -p 8080:8080 sopo
```

> Selenium 폴백을 쓰려면 `chromium`, `chromium-driver` 가 필요합니다 (`packages.txt` / `Dockerfile` 참고).

---

## 2. 지원 플랫폼

| 플랫폼 | 입력 형식 | 배송업체 | 통화 | 적용환율 |
|---|---|---|---|---|
| 쇼피 (Shopee) | 소포수령증 PDF | 두라로지스틱스 | MYR·PHP·SGD·THB·TWD·VND·IDR·BRL·MXN | 거래별 **발행일 일별** 매매기준율 |
| 라자다 (Lazada) | 소포수령증 PDF / 주문내역 xlsx | 용성종합물류 | MYR·PHP·SGD·VND·IDR | PDF=기간평균 / Excel=`deliveredDate` **일별** |
| 큐텐재팬 (Qoo10) | 소포수령증 PDF (+ 수동입력) | 국제로지스틱 | JPY | 거래기간이 속한 **반기말(6월/12월) 공식 월평균** |
| 이베이 / **린코스** | 소포수령증 PDF | 린코스(주) | USD·EUR·GBP·CAD·AUD 등 **혼합** | 발행월 **공식 월평균** |
| **Joom** | 상품 수령 및 운송 확인증 PDF | 에이치3네트웍스(H3 NETWORKS) | USD | 발송날짜 **일별** |
| **쇼피파이 (Shopify)** | `<스토어> orders *.csv` / `.xlsx` | (자체배송) | USD | `Fulfilled at` **일별** |

각 플랫폼의 상세 포맷·판별 규칙·계산 규칙은 [`docs/PLATFORMS.md`](docs/PLATFORMS.md) 를 보세요.

---

## 3. 프로젝트 구조

```
sopo/
├── app.py                          # Streamlit UI (업로드 → 파싱 → 환율 → 엑셀 생성 → 다운로드)
├── process.py                      # CLI 배치 진입점 (input/ → output/)
├── requirements.txt / packages.txt / Dockerfile
├── forms/                          # 관공서 제출 양식 원본 xlsx
│   ├── 수출실적명세서 양식.xlsx
│   └── 영세율첨부서류제출명세서 양식.xlsx
├── data/                           # 실행 중 생성되는 환율 캐시 (git 미추적)
│   ├── exchange_rate_cache.csv          # 일별 매매기준율
│   └── monthly_exchange_rate_cache.csv  # 공식 월평균 매매기준율
├── docs/
│   ├── PLATFORMS.md                # 플랫폼별 입력/판별/환율 규칙
│   └── ARCHITECTURE.md             # 데이터 흐름, 결과 dict 스키마, 신규 플랫폼 추가 절차
└── modules/
    ├── pdf_parser.py               # PDF 종류 판별 + 쇼피/라자다/큐텐/린코스(이베이)/Joom 파서
    ├── lazada_order_parser.py      # 라자다 주문내역 xlsx 파서
    ├── shopify_parser.py           # 쇼피파이 orders CSV/xlsx 파서
    ├── exchange_rate.py            # SMBS 일별·월평균 매매기준율 수집/캐시/조회
    ├── excel_writer.py             # 매출집계 워크북 생성 (총집계·월별집계·통화별·원본시트·환율시트)
    └── extra_docs.py               # 영세율첨부서류제출명세서 / 수출실적명세서 생성
```

---

## 4. 처리 흐름

```
업로드 파일
   │
   ├─ detect_pdf_type() / is_shopify_orders_file() / is_lazada_order_excel()
   │        ↓ 플랫폼 판별 (파일명 → 본문 표식 → 사용자 선택)
   │
   ├─ parse_pdf() / parse_shopify_orders() / parse_lazada_order_excel()
   │        ↓ 플랫폼별 result dict
   │
   ├─ 환율 수집
   │    ├─ fetch_all_currencies_for_period()          # 일별 (쇼피·라자다·Joom·쇼피파이)
   │    └─ fetch_monthly_avg_currencies_for_period()  # 공식 월평균 (린코스·큐텐)
   │
   ├─ generate_excel()          → 매출집계_*.xlsx
   └─ build_declaration_rows()  → 영세율첨부서류제출명세서_*.xlsx / 수출실적명세서_*.xlsx
```

---

## 5. 환율 규칙 (중요)

* **일별 매매기준율** — 해당 일자가 휴일이면 **직전 영업일** 환율을 사용합니다.
  (연초 등을 대비해 신고기간 시작일보다 `RATE_LOOKBACK_DAYS=7`일 앞에서부터 수집합니다.)
* **월평균 매매기준율** — 서울외국환중개 `MonAvgStdExRate.jsp` 의 **공식 값만** 사용합니다.
  일별 환율을 자체 평균내어 대체하지 않으며, 공식 값을 못 얻으면 생성이 중단됩니다.
* **JPY·IDR·VND** 는 SMBS가 100통화 단위로 고시합니다.
  내부 계산은 **1통화 단위**(소수점 4자리)로 정규화하고, `환율(통화)` 시트에만 원문과 같은 100통화 단위로 표시합니다.

---

## 6. 산출물

* `매출집계_<기간>.xlsx`
  * `총집계` — 플랫폼·통화별 외화/원화 합계
  * `월별집계` — 월 × 구분 × 통화 집계
  * `USD` / `EUR` / `MYR` … — 통화별 수출신고 프로그램 업로드용 시트
  * `쇼피(MYR)` / `라자다(주문내역)` / `이베이` / `Joom` / `쇼피파이(스토어명)` — 원본 자료 시트
  * `환율(USD)` … — 적용한 매매기준율 근거 시트
* `영세율첨부서류제출명세서_<업체>_<전체|YYYY년MM월>.xlsx`
* `수출실적명세서_<업체>.xlsx`

---

## 7. 접근 권한

`.streamlit/secrets.toml` 에 `[auth]` 섹션이 있으면 Google 로그인이 활성화되고,
`app.py` 의 `ALLOWED_EMAILS` 목록에 있는 계정만 사용할 수 있습니다.
`[auth]` 가 없으면 로그인 없이 동작합니다.
