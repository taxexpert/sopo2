# 📦 소포수령증 자동화 (sopo)

해외 오픈마켓·배송대행사가 발급한 **해외배송 소포수령증(PDF)** 과 **주문내역(Excel/CSV)** 을 읽어
부가가치세 신고용 **매출집계 / 영세율첨부서류제출명세서 / 수출실적명세서** 를 자동 생성합니다.

환율은 **서울외국환중개(SMBS)** 에서 직접 수집합니다.

> ### 🔗 사용하실 주소는 여기입니다 → **https://soposopo.streamlit.app/**
> 설치할 것도, 내려받을 것도 없습니다. 브라우저에서 위 주소만 열면 됩니다.

---

## 이 문서는 두 가지 역할로 나눠져 있습니다

| | 대상 | 읽을 곳 |
|---|---|---|
| 🙋 **사용하는 분** | 실제로 신고서류를 만드는 담당자 | [A. 사용 안내](#a-사용-안내-담당자용) |
| 🛠️ **관리하는 분** | 앱을 고치고 배포하는 담당자 | [B. 운영·개발 안내](#b-운영개발-안내-관리자용) |

---

# A. 사용 안내 (담당자용)

## A-1. 접속하기

1. 브라우저에서 **https://soposopo.streamlit.app/** 를 엽니다. (즐겨찾기 해두시면 편합니다)
2. **로그인 화면이 뜨면** `Google 계정으로 로그인` 을 누르고 **회사 구글 계정**으로 로그인합니다.
   * 미리 등록된 계정만 들어갈 수 있습니다. `접근 권한이 없는 계정입니다.` 가 뜨면 관리자에게
     사용하실 구글 계정 주소를 알려주고 등록을 요청하세요.
   * 개인 지메일로 로그인되어 있으면 권한 오류가 날 수 있습니다. 좌측 `로그아웃` 후 회사 계정으로 다시 로그인하세요.
3. 아무도 한동안 쓰지 않으면 앱이 **잠자기 상태**가 됩니다.
   `Yes, get this app back up!` 같은 버튼이 보이면 눌러주세요. 1~2분 뒤 정상 화면이 뜹니다.

> 💡 **설치는 필요 없습니다.** 아래 B장에 나오는 `pip install`, `streamlit run` 같은 명령은
> 관리자가 앱을 고칠 때 쓰는 것이고, 담당자는 위 주소만 쓰시면 됩니다.

## A-2. 화면 순서대로 따라하기

앱 화면은 STEP 1 → 4 순서로 되어 있고, 생성이 끝나면 아래에 STEP 5 내려받기가 나타납니다.

**STEP 1 — 수령증·주문내역 올리기**
소포수령증 PDF와 주문내역 Excel/CSV를 한꺼번에 올립니다. 여러 플랫폼을 섞어서 올려도 됩니다.
파일명·본문을 보고 플랫폼을 자동으로 알아내며, 못 알아낸 파일만 직접 골라주면 됩니다.

> ⚠️ **올릴 수 있는 자료는 플랫폼이 발행한 원본 서식뿐입니다.** 쇼피·라자다·큐텐재팬·이베이(린코스)·Joom 이
> 발행한 소포수령증 PDF, 라자다 주문내역 xlsx, 쇼피파이 orders 파일만 읽습니다.
> 직접 만든 집계표, 서식이 다른 자료, 화면을 캡처한 이미지는 인식하지 못합니다. 지원 목록은 A-4 를 보세요.

**STEP 2 — 큐텐재팬 정보 확인**
큐텐재팬 PDF를 올리면 자동으로 목록에 채워지고, `큐텐재팬 입력 목록 (n건)` 이 펼쳐집니다.
빠진 건은 이 목록에서 직접 추가합니다. 큐텐 자료가 없으면 접힌 채로 두고 넘어가세요.

**STEP 3 — 만들 문서 고르기**
매출집계 / 영세율첨부서류제출명세서 / 수출실적명세서 중 필요한 것을 체크합니다.

**STEP 4 — 생성**
`엑셀 파일 생성` 을 누르면 파일을 읽고 → 환율을 조회하고 → 엑셀을 만듭니다.
진행 상황이 검은 로그창에 그대로 표시되고, 끝나면 **STEP 5 내려받기**에 다운로드 버튼이 생깁니다.

> ⏳ 환율은 서울외국환중개에서 그때그때 받아오므로, 처음 조회하는 기간은 시간이 조금 걸립니다.
> 로그가 멈춘 것처럼 보여도 창을 닫지 마세요.

## A-3. 받게 되는 파일

| 파일 | 내용 |
|---|---|
| `매출집계_<기간>.xlsx` | `총집계`(플랫폼·통화별 합계), `월별집계`, 통화별 시트(수출신고 프로그램 업로드용), 플랫폼별 원본 시트, `환율(통화)` 근거 시트 |
| `영세율첨부서류제출명세서_<업체>_<기간>.xlsx` | 제출용 양식 |
| `수출실적명세서_<업체>.xlsx` | 제출용 양식 |

## A-4. 지원하는 플랫폼

| 플랫폼 | 올리는 파일 | 배송업체 | 통화 | 적용환율 |
|---|---|---|---|---|
| 쇼피 (Shopee) | 소포수령증 PDF | 두라로지스틱스 | MYR·PHP·SGD·THB·TWD·VND·IDR·BRL·MXN | 거래별 **발행일 일별** 매매기준율 |
| 라자다 (Lazada) | 소포수령증 PDF / 주문내역 xlsx | 용성종합물류 | MYR·PHP·SGD·VND·IDR | PDF=기간평균 / Excel=`deliveredDate` **일별** |
| 큐텐재팬 (Qoo10) | 소포수령증 PDF (+ 화면 입력) | 국제로지스틱 | JPY | 거래기간이 속한 **반기말(6월/12월) 공식 월평균** |
| 이베이 / 린코스 | 소포수령증 PDF | 린코스(주) | USD·EUR·GBP·CAD·AUD 등 **혼합** | 발행월 **공식 월평균** |
| Joom | 상품 수령 및 운송 확인증 PDF | 에이치3네트웍스(H3 NETWORKS) | USD | 발송날짜 **일별** |
| 쇼피파이 (Shopify) | `<스토어> orders *.csv` / `.xlsx` | (자체배송) | USD | `Fulfilled at` **일별** |

표에 없는 플랫폼, 그리고 표에 있어도 **발행 서식이 아닌 자료**(직접 만든 집계표, 결제·정산 내역, 화면 캡처)는 처리할 수 없습니다.
파일명을 바꿔도 결과는 같습니다. 판별은 PDF 본문 문구와 Excel/CSV 열 구성으로 합니다.

플랫폼별 상세 규칙은 [`docs/PLATFORMS.md`](docs/PLATFORMS.md) 참고.

## A-5. 자주 겪는 상황

| 증상 | 해결 |
|---|---|
| `접근 권한이 없는 계정입니다.` | 관리자에게 구글 계정 등록 요청 (A-1 참고) |
| 화면이 안 뜨고 깨우기 버튼만 보임 | 버튼을 누르고 1~2분 대기 |
| 파일이 `미확인` 으로 잡힘 | 화면에서 플랫폼을 직접 선택 후 진행 (발행 서식이 아니면 선택해도 값이 나오지 않습니다) |
| 직접 만든 집계표·서식이 다른 파일을 올려 값이 안 나옴 | 플랫폼에서 내려받은 원본 서식으로 다시 올리기 (A-4 참고) |
| 처리 중 오류 로그가 뜸 | 검은 로그창 내용을 **그대로 복사**해서 관리자에게 전달 |
| 금액·건수가 원본과 안 맞음 | 원본 자료와 함께 관리자에게 전달 ([`docs/VERIFICATION.md`](docs/VERIFICATION.md) 절차로 대사) |

> ⚠️ 업로드한 파일에는 사업자번호·구매자 정보가 들어 있습니다. 개인 계정으로 로그인해서 쓰거나,
> 산출물을 외부에 공유하지 마세요.

---

# B. 운영·개발 안내 (관리자용)

## B-1. 배포 구조

```
GitHub  taxexpert/sopo2  (main 브랜치)
   │  push 하면 자동 재배포
   ▼
Streamlit Community Cloud   ← 계정: taxexpert
   │  entrypoint: app.py
   ▼
https://soposopo.streamlit.app/     ← 담당자들이 쓰는 주소
```

**담당자는 이 주소만 쓰고, 로컬 실행은 관리자가 고칠 때만 씁니다.**
`main` 에 push 하는 순간 운영 사이트가 바뀌므로, 검증 후 push 하세요.

Streamlit Cloud 가 참고하는 파일:

| 파일 | 역할 |
|---|---|
| `app.py` | 앱 진입점 (Cloud 설정의 Main file path) |
| `requirements.txt` | 파이썬 패키지 |
| `packages.txt` | OS 패키지 (`chromium`, `chromium-driver` — 환율 Selenium 폴백용) |
| `Dockerfile` | Streamlit Cloud 는 쓰지 않음. Cloud Run 등 다른 곳에 올릴 때만 사용 |

## B-2. 로그인·권한 설정

로그인은 **켜고 끌 수 있는 선택 기능**입니다. — [`app.py:247`](app.py:247)

* Secrets 에 `[auth]` 섹션이 **있으면** Google 로그인이 켜집니다.
* **없으면** 로그인 없이 누구나 접속합니다. → 운영 주소에서는 반드시 `[auth]` 를 설정하세요.

Streamlit Cloud 에서는 `.streamlit/secrets.toml` 를 커밋하지 않고
(→ `.gitignore` 로 제외됨) **앱 대시보드의 Settings → Secrets** 패널에 붙여넣습니다.

```toml
[auth]
redirect_uri = "https://soposopo.streamlit.app/oauth2callback"
cookie_secret = "<임의의 긴 문자열>"
client_id = "<Google OAuth 클라이언트 ID>"
client_secret = "<Google OAuth 클라이언트 시크릿>"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

> Google Cloud Console 의 OAuth 클라이언트에도 위 `redirect_uri` 가 승인된 리디렉션 URI 로
> 등록되어 있어야 합니다.

**사용자 추가/삭제** 는 [`app.py:52`](app.py:52) 의 `ALLOWED_EMAILS` 목록을 고치고 `main` 에 push 하면 됩니다.
(로그인은 통과해도 이 목록에 없으면 차단됩니다.)

로컬에서 로그인까지 테스트하려면 같은 내용을 `.streamlit/secrets.toml` 로 만들고
`redirect_uri` 만 `http://localhost:8501/oauth2callback` 로 바꿉니다.

## B-3. 로컬 실행 (개발용)

```bash
python -m pip install -r requirements.txt
```

```bash
python -m streamlit run app.py
```

CLI 배치 (`input/` 폴더 일괄 처리 → `output/`):

```bash
python process.py --year 2026 --month 03
```

Docker (Streamlit Cloud 외 환경에 올릴 때):

```bash
docker build -t sopo . && docker run -p 8080:8080 sopo
```

## B-4. 환율 캐시 주의

수집한 환율은 `data/` 아래 CSV 캐시에 저장해 같은 구간을 반복 조회하지 않습니다.
다만 **Streamlit Cloud 의 디스크는 휘발성**이라 앱이 재시작(재배포·잠자기 해제)되면 캐시가 비고,
그 뒤 첫 조회는 다시 SMBS 를 호출해 느려집니다. 정상 동작이며 결과값에는 영향이 없습니다.

## B-5. 환율 규칙 (중요)

* **일별 매매기준율** — 해당 일자가 휴일이면 **직전 영업일** 환율을 사용합니다.
  (연초 등을 대비해 신고기간 시작일보다 `RATE_LOOKBACK_DAYS=7`일 앞에서부터 수집합니다.)
* **월평균 매매기준율** — 서울외국환중개 `MonAvgStdExRate.jsp` 의 **공식 값만** 사용합니다.
  일별 환율을 자체 평균내어 대체하지 않으며, 공식 값을 못 얻으면 생성이 중단됩니다.
* **JPY·IDR·VND** 는 SMBS 가 100통화 단위로 고시합니다.
  내부 계산은 **1통화 단위**(소수점 4자리)로 정규화하고, `환율(통화)` 시트에만 원문과 같은 100통화 단위로 표시합니다.

## B-6. 프로젝트 구조

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
├── samples/                        # 회귀 검증용 예시 자료 (고객 파일은 git 미추적)
├── docs/
│   ├── PLATFORMS.md                # 플랫폼별 입력/판별/환율 규칙
│   ├── ARCHITECTURE.md             # 데이터 흐름, 결과 dict 스키마, 신규 플랫폼 추가 절차
│   └── VERIFICATION.md             # 산출물 대사·검증 절차
└── modules/
    ├── pdf_parser.py               # PDF 종류 판별 + 쇼피/라자다/큐텐/린코스(이베이)/Joom 파서
    ├── lazada_order_parser.py      # 라자다 주문내역 xlsx 파서
    ├── shopify_parser.py           # 쇼피파이 orders CSV/xlsx 파서
    ├── exchange_rate.py            # SMBS 일별·월평균 매매기준율 수집/캐시/조회
    ├── excel_writer.py             # 매출집계 워크북 생성 (총집계·월별집계·통화별·원본시트·환율시트)
    └── extra_docs.py               # 영세율첨부서류제출명세서 / 수출실적명세서 생성
```

## B-7. 처리 흐름

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

신규 플랫폼 추가 절차는 [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) 를 보세요.
