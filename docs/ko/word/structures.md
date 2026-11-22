# 1. 2D 구조식 · 반응식 작성

오프라인 **Ketcher 2D**로 Word 본문이나 표 셀에 인라인 구조식과 다단계 합성 반응식을 그리거나, 문서 속 SMILES·InChI·화합물명을 **`Paste Structure`** 로 `0.1초` 만에 그림으로 변환하고, 언제든 **더블클릭**해 레이아웃 깨짐 없이 제자리에서 수정합니다.

## 1. 상황별 가장 빠른 입력 방법

| 시작 상태 | 실행할 동작 | Word에서의 결과 |
| :--- | :--- | :--- |
| **빈 커서 위치** (새 분자·반응식 그리기) | **`[SMILES]` ➔ `[Insert / Edit Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Insert / Edit Structure...`**) | 오프라인 **Ketcher 2D** 실행 후 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> 시 비율 고정 PNG 삽입 |
| **본문/표에 적힌 텍스트** (`CC(=O)Oc1ccccc1C(=O)O`, `InChI=...`, `Aspirin`, `50-78-2`) | 텍스트를 드래그 선택 ➔ **`[SMILES]` ➔ `[Paste Structure]`** | 선택한 텍스트 자리가 `0.1초` 만에 인라인 구조식 그림으로 **제자리 치환** |
| **클립보드에 복사한 텍스트** (PubChem, Excel, ChemDraw, ELN) | 텍스트 선택 없이 커서만 두고 **`[Paste Structure]`** | 클립보드의 SMILES·InChI·KET·화합물명을 커서 위치에 즉시 삽입 |
| **적어 둔 분자를 살짝 수정해서 넣고 싶을 때** | 본문의 `Aspirin`이나 SMILES를 드래그 선택 ➔ **`[Insert / Edit Structure]`** | **Ketcher 2D 캔버스에 해당 구조가 미리 그려진 상태로 열려** 바로 치환기 편집 가능 |

::: tip 숨은 꿀팁: 본문 텍스트를 Ketcher에 바로 불러오기
굳이 먼저 그림으로 바꿨다가 다시 더블클릭할 필요가 없습니다. 원고에 SMILES나 `Aspirin` 같은 물질명이 적혀 있다면 **글자를 드래그한 상태에서 `[Insert / Edit Structure]`** 를 누르세요. 해당 분자가 Ketcher 캔버스에 자동으로 로드된 채 열립니다.
:::

## 2. 인라인 구조식 그리기 & `0.1초` 텍스트 변환

### 본문 커서 위치에 직접 그리기 (`Ketcher 2D`)

<FeatureMedia id="word-ch1-insert-ketcher" />

1. 구조식을 넣을 문단이나 표 셀 안에 커서를 둡니다.
2. 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** 를 클릭해 **Ketcher 2D**에서 분자를 그립니다(단축키 및 배경 옵션은 [Ketcher 2D 상세 매뉴얼](/ko/editor/) 참고).
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**)를 누르면 가로세로 비율이 잠긴(`LockAspectRatio`) 고해상도 인라인 그림(`InlineShape`)으로 삽입됩니다.

### SMILES · InChI · 화합물명 `0.1초` 즉시 변환 (`Paste Structure`)

<FeatureMedia id="word-ch2-quick-insert" />

* **SMILES · InChI · KET · 물질명/CAS 모두 지원:** 자주 쓰는 의약품·용매·CAS 번호(`Aspirin`, `Caffeine`, `Paracetamol`, `THF`, `50-78-2`)는 내장 사전으로 **0ms 오프라인 변환**되며, 사전에 없는 영문 물질명은 PubChem에서 1회 조회 후 디스크에 캐시됩니다.
* **선택 영역 즉시 치환:** 본문 문장이나 표 셀 안의 텍스트를 드래그해 누르면 원래 글자를 지우고 그 자리에 깔끔하게 구조식을 끼워 넣습니다.

### 인라인 그림(`InlineShape`) vs. 부동 그림(`Shape`)

* **인라인 그림(기본값):** 텍스트 줄의 일부로 취급되어 앞 문단을 고쳐도 구조식 위치가 밀리거나 페이지 밖으로 튀지 않으므로 논문 본문과 표 셀에 가장 적합합니다.
* **부동 그림(텍스트 배치 변경 시):** 다단 Figure나 콜아웃 박스 배치를 위해 *사각형/텍스트 앞*으로 바꾸더라도 `func_SMILES2IMG`가 더블클릭 이벤트를 가로채므로 Word 그림 서식 창 대신 **Ketcher 2D가 제자리에서 열립니다**.

## 3. 다단계 합성 반응식 (`KET`) 무손실 삽입

반응 화살표(`➔`), `+` 기호, 시약·촉매·온도·수율 주석이 포함된 합성 경로를 언제든 다시 고칠 수 있게 `.docx` 원고에 저장합니다:

<FeatureMedia id="word-ch4-reaction-scheme" />

1. 반응식을 배치할 문단이나 넓은 Figure 셀에 커서를 두고 **`[SMILES]` ➔ `[Insert / Edit Structure]`** 를 클릭합니다.
2. **Ketcher 2D**에서 반응물, 화살표(`➔`), `+` 기호, 텍스트 도구(`T`로 `H₂SO₄, 85 °C, 92%` 등 주석 입력)를 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (**`Apply to Document`**)를 누릅니다.
4. **`.docx` 무손실 좌표 보존:** 화살표·`+`·텍스트가 포함된 반응식은 전체 KET JSON 좌표가 그림의 **대체 텍스트(Alt Text)** (`KET:...`)에, 정규 SMILES가 `Title`에 함께 저장됩니다. 파일을 저장하고 공동저자와 주고받은 뒤 다시 열어 **더블클릭**해도 화살표 각도와 텍스트 배치 그대로 복원됩니다.

## 4. 더블클릭 제자리 수정 & 다른 앱에서 가져오기

* **크기·줄 배치 유지:** 인라인 그림이든 부동 그림이든 **더블클릭**하면 `0.05초` 만에 Ketcher 2D가 열리며, 수정 후 적용 시 기존 너비와 문단 배치를 유지한 채 교체됩니다.
* **Excel ➔ Word 옮기기:** Excel 시트의 구조식 셀(`=SMILES2IMG(...)` 또는 `=KET2IMG(...)`)에서 **`[Copy KET]`**(반응식) 또는 **`[Copy SMILES]`** 를 누른 뒤, Word에서 **`[Paste Structure]`** 를 클릭하면 곧바로 동일한 구조식이 들어갑니다.
* **ChemDraw ➔ Word 전환:** 기존 ChemDraw 파일의 구조나 반응식을 선택해 **Edit ➔ Copy As ➔ SMILES**로 복사한 뒤 Word에서 **`[Paste Structure]`** 를 실행하면 더블클릭 편집 가능한 `func_SMILES2IMG` 개체로 전환됩니다.

## 관련 가이드

- [**2. 화합물 표(SAR) & 실험 물성(SI)**](/ko/word/tables-properties) — 표 셀 위쪽 잘림 방지 설정 및 오타 없는 화학식·분자량·HRMS 복사
- [**Ketcher 2D 편집기 상세 매뉴얼**](/ko/editor/) — 고리 템플릿, 단축키, 배경 옵션 전체 안내
- [**더블클릭 · 재편집 문제 해결**](/ko/word/troubleshooting/double-click) — 더블클릭 시 그림 서식 창이 열릴 때 확인 사항
