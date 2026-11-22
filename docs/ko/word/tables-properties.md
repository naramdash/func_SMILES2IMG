# 2. 화합물 표(SAR) & 실험 물성(SI)

Word 논문용 SAR·화합물 요약 표의 열 너비 틀어짐과 그림 잘림을 원천 차단하고, **`Molecule Info`** 로 화학식·분자량·HRMS 정밀질량·LogP·Lipinski 판정을 오타 없이 즉시 복사해 채웁니다.

<FeatureMedia id="word-ch3-compound-table" />

## 1. 깨지지도 잘리지도 않는 Word 화합물 표 3원칙

Word 기본 표 설정에서는 구조식을 넣을 때 열이 찌그러지거나 그림 윗부분이 잘리는 일이 흔합니다. 표를 만든 직후 아래 3가지만 맞혀 두세요:

1. **열 너비 고정:** **표 우클릭 ➔ 자동 맞춤 ➔ 열 너비 고정**을 설정합니다. 이후 구조식을 더블클릭해 작용기를 늘려도 옆의 물성·IC₅₀ 열 너비가 임의로 좁아지지 않습니다.
2. **문단 줄 간격을 `배수` 또는 `1줄`로 유지 (그림 위쪽 잘림 방지!):** 표 셀의 줄 간격이 **`고정값`(예: *고정값 12 pt*)** 으로 되어 있으면 인라인 구조식 상단이 칼로 자른 듯 잘려 보입니다. 구조식이 들어갈 열을 선택하고 **홈 ➔ 줄 및 단락 간격 ➔ 줄 간격 옵션**에서 **`1줄`** 또는 **`배수`(예: `1.15`)** 로 변경하세요.
3. **음영 셀 배경 맞춤:** 표 행에 회색이나 컬러 음영이 들어가 있다면 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Standard Color (Transparent)]`** 로 투명하게 녹이거나 **`[White Background (Color)]`** 로 흰 배경 프레임을 씌워 시인성을 높입니다.

::: tip 꿀팁: SMILES/물질명이 적힌 표 열을 한 행씩 빠르게 구조식으로 바꾸기
이미 Word 표에 `Aspirin`, `Caffeine`, `50-78-2` 또는 SMILES가 적힌 열이 있다면 텍스트를 드래그하고 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 클릭해 `0.1초` 만에 비율 고정 구조식으로 치환할 수 있습니다.
:::

## 2. `Molecule Info`로 표 · Experimental Section 오타 없이 채우기

Word의 구조식 그림을 선택하거나—**본문에 적힌 SMILES/화합물명 텍스트만 드래그한 상태에서**—상단 **`[SMILES]` ➔ `[Molecule Info]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Copy Molecule Info`**)를 클릭하면 로컬에서 즉시 계산된 17종 물성을 원하는 서식으로 복사할 수 있습니다:

<FeatureMedia id="common-molecule-info" />

| Word에서 작성 중인 항목 | 선택할 `Molecule Info` 메뉴 | 붙여넣기(<kbd>Ctrl</kbd> + <kbd>V</kbd>) 결과 |
| :--- | :--- | :--- |
| **새 17열 물성 요약 표 생성** | **`Copy All Summary (Text)`** | 17개 물성 머리글(Header) + 데이터 행이 서식 있는 Word 표(HTML/TSV)로 즉시 생성됨 |
| **기존 표의 빈 행에 값 채우기** | **`Copy as Excel Row (TSV Columns)`** | 헤더 없이 17개 탭 구분 값만 복사되어 기존 표 행 셀에 맞추어 붙여넣기 가능 |
| **Experimental Section · HRMS 기재** | **`Copy Chemical Formula`** / **`Copy Molecular Weight (MW)`** / **`Detailed Molecule Info Window...`** | Hill 화학식(`C9H8O4`), 평균 분자량(`180.16`), HRMS 대조용 단일동위원소 정밀질량(`180.0423 Da`) 개별 복사 |
| **Supporting Information (SI) 요약 목록** | **`Copy as Multi-Line List`** | 물질명, IUPAC명, 물성 12종, SMILES, InChI, InChIKey가 정리된 줄바꿈 목록 복사 |

::: tip 숨은 꿀팁: 그림을 넣지 않고 텍스트에서도 바로 물성·SVG 추출
구조식 그림을 아직 삽입하지 않았더라도 본문의 SMILES 문자열이나 물질명(`Ibuprofen`)을 드래그한 상태에서 **`[Molecule Info]`** 또는 **`[Copy SVG]`** 를 실행하면 즉시 계산·복사됩니다.
:::

## 3. Excel SAR 시트와 연동해 표 만들기

이미 Excel에서 `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2LIPINSKI`, `=SMILES.SIMILARITY` 등으로 스크리닝 표를 정리해 두었다면:

* **물성 데이터 일괄 붙여넣기:** Excel에서 숫자·텍스트 물성 행을 복사해 Word 표에 붙여넣은 뒤 **열 너비 고정**과 줄 간격 **`1줄`** 을 확인하고, 구조식 열에만 **`[Paste Structure]`** 나 **`[Insert / Edit Structure]`** 로 2D 그림을 채우세요.
* **수정 시 즉시 동기화:** Word 표에서 구조식을 **더블클릭**해 치환기를 바꿨다면 바로 **`[Molecule Info]`** 를 눌러 변경된 화학식과 정밀질량을 확인·갱신할 수 있습니다.

## 관련 가이드

- [**스타일 & 출력 상세 매뉴얼**](/ko/editor/publishing) — `Molecule Info` 5가지 클립보드 복사 서식(Summary, TSV, Multi-Line, Markdown, JSON) 실제 출력 예시
- [**4. 투고 마감 · 벡터 SVG · 공동저자 공유**](/ko/word/publishing) — 투고 직전 표와 본문의 모든 구조식을 클릭 한 번으로 ACS 흑백 통일하기
- [**표 잘림 · 일괄 스타일 문제 해결**](/ko/word/troubleshooting/table-layout) — 표 셀 상단 그림 잘림, 줄 간격 `고정값` 해제, `Batch Style` 범위 문제 해결
