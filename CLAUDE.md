# CLAUDE.md

소포수령증 자동화 Streamlit 앱. 해외 마켓(쇼피·라자다·큐텐재팬·이베이/린코스·Joom·쇼피파이)의
소포수령증 PDF와 주문내역 파일을 파싱해 매출집계·영세율첨부서류제출명세서·수출실적명세서를 생성한다.

**신고서류 동작을 고치기 전에 `docs/HANDOFF_V54.md`(전임자 확정 규칙·금지 회귀 14종)와
`docs/POLICY.md`(사용자 확정 정책)를 먼저 확인한다.** 검증에서 "오류"로 보여도 이 두 문서에
있는 동작(발급일자=선적일자, L/C칸 운송장, 기타건수 1, 큐텐 반기말 등)은 의도된 것이다.

## PR을 만들 때 — base 저장소를 반드시 명시할 것

현재 저장소는 **`taxexpert/sopo2`**다 (구 `taxexpert/sopo`에서 이전). 그리고 이전 담당자의
**`rambbo1/sopo`를 fork한 저장소**라, GitHub이 PR 생성 화면에서 base 저장소를 부모인
`rambbo1/sopo`로 **기본 지정한다**. 명시하지 않으면 회사 저장소가 아니라 이전 담당자 저장소로
PR이 올라간다.

```bash
gh pr create --repo taxexpert/sopo2 --base main --head <branch>
```

웹에서 만들 때는 저장소 내부 비교 URL을 쓴다. base가 부모로 갈 수 없다.

```
https://github.com/taxexpert/sopo2/compare/main...<branch>?expand=1
```

`https://github.com/taxexpert/sopo2/pull/new/<branch>` 형태는 부모 저장소로 기본 지정되므로 쓰지 않는다.
`git push` 직후 원격이 출력하는 링크가 바로 이 형태다.

PR 화면 맨 위 드롭다운 네 개 중 **첫 번째 `base repository`**가 요주의 지점이다.
화면에 `rambbo1`이 보이면 잘못된 것이고, 양쪽 모두 `taxexpert/sopo2`여야 한다.

## 배포

`main`에 push하면 Streamlit Cloud가 자동 재배포한다 (https://soposopo.streamlit.app/).
따라서 **작업 브랜치를 머지하기 전에는 배포 화면에 반영되지 않는다.** 브랜치에서 확인하려면 로컬로 실행한다.

```bash
python -m streamlit run app.py
```

## 작업 환경 주의

- 이 Windows 머신에는 `gh` CLI와 `node`가 설치돼 있지 않다. `gh`가 필요하면
  `winget install --id GitHub.cli -e` 후 터미널을 재시작해야 PATH에 잡힌다.
- 셸은 PowerShell이 기본이고 Bash도 쓸 수 있다. 각자 문법이 다르다.

## 구조

- `app.py` — Streamlit UI, 업로드·판별·오케스트레이션. 화면 스타일(CSS 토큰·헤더·단계 머리)도 여기 상단에 있다
- `modules/` — 플랫폼별 파서, 환율 수집(서울외국환중개), 엑셀 생성
- `forms/` · `samples/` — 서식 템플릿과 샘플 입력
- `process.py` — CLI 경로

## 화면 규칙

사용자 화면은 세무 서식의 시각 언어를 따른다. 새 UI를 추가할 때 이 규칙을 깨지 않는다.

- 이모지를 아이콘으로 쓰지 않는다. 단계는 번호·세로 괘선·제목으로 표기한다
  (처리 로그 상자 안의 상태 글리프는 예외)
- 구획은 `st.divider()`가 아니라 단계 머리가 갖는 상단 괘선으로 나눈다
- 색은 엑스퍼트 로고에서 온 브랜드 블루와 하늘색뿐이다. CSS 변수(`--accent`, `--sky`,
  `--ink`, `--line`, `--surface`)를 쓰고 색상값을 직접 박지 않는다
- 채워진 버튼은 `--accent-solid`/`--on-accent` 쌍을 쓴다. `--accent`를 배경에 그대로 쓰면
  어두운 테마에서 대비가 깨진다
- 사용자 입력(파일명 등)을 `unsafe_allow_html`로 출력할 때는 `html.escape`를 거친다
