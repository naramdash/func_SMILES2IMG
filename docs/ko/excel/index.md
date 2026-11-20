# Excel 개요

Microsoft Excel에서 `func_SMILES2IMG`를 사용하면 간단한 셀 수식으로 수백 행의 2D 화학 구조식을 렌더링하고, 분자 물성·Lipinski 약물성·SMARTS 부분구조 검색을 로컬에서 즉시 계산하며, `.sdf` 화학 데이터베이스를 원클릭으로 가져오고 내보낼 수 있습니다.

<FeatureMedia id="excel-overview" />

::: tip 10개 탭 자체 완결형 Excel Playbook 다운로드
웹 문서 없이 엑셀 파일 하나만으로 40개 전체 함수, 실전 합성 당량·Hit-to-Lead 템플릿, 오프라인 함수 사전까지 바로 실습해 보세요:  
<a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx</strong></a> · [10개 탭 상세 구성 및 미리보기 →](/ko/excel/playbook)
:::

## Excel에서 구조식을 다루는 3가지 방법

| 방법 | 사용법 | 추천 용도 |
| :--- | :--- | :--- |
| **셀 수식 입력** | `=SMILES2IMG(A2)`, `=NAME2IMG("Aspirin")`, `=KET2IMG(A2)` | 수백 행 일괄 렌더링, 정렬·필터 연동 |
| **상단 `[SMILES]` 리본** | **`[Insert / Edit (SMILES)]`** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>), **`[Paste Structure]`** | 분자/반응식 직접 그리기, 클립보드 구조 즉시 붙여넣기 |
| **마우스 조작** | 구조식 셀 **더블클릭** (제자리 편집) / **우클릭** ➔ `Chemical Structure` | 빠른 수정, SVG 복사, Molecule Info 확인 |

::: info Microsoft 365 / Excel 2024+ vs. Excel 2016~2021 (`.FLOAT`)
* **Microsoft 365 / Excel 2024 이상:** `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)`로 정렬·필터와 연동되는 네이티브 **인셀 이미지**를 생성합니다.
* **Excel 2016 / 2019 / 2021:** 뒤에 **`.FLOAT`**(`=SMILES2IMG.FLOAT(A2)`)를 붙여 셀에 고정되는 그림 도형(`xlMoveAndSize`)으로 생성합니다. 자세한 내용은 [구버전 Excel (.FLOAT)](/ko/excel/legacy-float)를 참고하세요.
:::

## 다음 단계

- [**구조식 표 · 이름/CAS**](/ko/excel/structures) — SMILES, 화합물 이름, 용매 약어, CAS 번호로 대량 구조식 표 만들기
- [**스타일 · 하이라이트 · 반응식**](/ko/excel/styles-reactions) — ACS 흑백, 다크 모드, SMARTS 하이라이트, NMR 번호, `=KET2IMG`
- [**물성 · Lipinski 약물성**](/ko/excel/properties) — 분자량, 정밀질량, LogP, TPSA, 고리 수, Fsp3, 키랄 중심 계산
- [**40개 함수 사전**](/ko/excel/formulas) — 전체 함수 인자 및 반환 형식 레퍼런스
