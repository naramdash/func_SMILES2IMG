# Excel 개요

Microsoft Excel에서 **`func_SMILES2IMG`** 를 사용하면 셀 수식으로 2D 화학 구조식을 렌더링하고, 분자 물성·LogP/TPSA·3D 복잡도·SMARTS 부분구조 검색을 로컬 오프라인으로 계산하며, `.sdf` 화학 데이터베이스를 수식 시트로 불러올 수 있습니다.

<FeatureMedia id="excel-overview" />

::: tip 10개 탭 자체 완결형 Excel Playbook 다운로드
웹 문서 없이 엑셀 파일 하나만으로 49개 전체 함수, 실전 합성 당량·Hit-to-Lead 템플릿, 오프라인 함수 사전까지 실습해 보세요:  
<a href="/func_SMILES2IMG/examples/func_SMILES2IMG_Excel_Playbook.xlsx" download="func_SMILES2IMG_Excel_Playbook.xlsx"><strong>func_SMILES2IMG_Excel_Playbook.xlsx</strong></a> · [10개 탭 구성 및 추천 실습 코스 →](/ko/excel/playbook)
:::

## 빠른 시작

**전체 복사** 를 클릭해 엑셀 **`A1`** 셀에 붙여넣거나 각 행을 클릭해 입력값과 수식을 함께 복사해 보세요 (**Excel 2016~2021**에서는 뒤에 `.FLOAT`를 붙입니다: `=SMILES2IMG.FLOAT(A2)`):

<FormulaTable
  title="빠른 시작 시트"
  :columns='["SMILES / 입력값", "수식"]'
  :rows='[
    { input: "CCO", formula: "=SMILES2IMG(A2)", desc: "A2 SMILES로 2D 구조식 렌더링", out: "에탄올 2D" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A3)", desc: "평균 분자량 로컬 계산", out: "180.16" },
    { input: "", formula: "=NAME2IMG(\"Aspirin\")", desc: "시약명·CAS 단독 입력 2D 구조식", out: "아스피린 2D" }
  ]'
/>

**행 높이**(`62~80 pt`)와 **열 너비**를 넓히면 가로세로 비율이 유지된 채 선명하게 확대됩니다. 구조식 셀을 **더블클릭**하면 오프라인 **Ketcher 2D** 편집기가 열려 제자리에서 수정할 수 있습니다.

## Excel에서 구조식을 다루는 3가지 방법

| 방법 | 사용법 | 추천 용도 |
| :--- | :--- | :--- |
| **셀 수식 입력** | `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=SMILES2MW(B2)` | 대량 행 일괄 렌더링 및 정렬·필터 자동 연동 |
| **상단 `[SMILES]` 리본** | **`[Insert / Edit]`**, **`[Paste Structure]`**, **`[3D Structure (Mol*)]`** | 분자·반응식 직접 그리기, 클립보드 구조 변환, 3D 단백질 삽입 |
| **마우스 조작** | 구조식 셀 **더블클릭** (제자리 편집) / **우클릭** ➔ `Chemical Structure` | 구조 수정, 벡터 SVG 복사, `Molecule Info` 물성 확인 |

::: info Microsoft 365 / Excel 2024+ vs. Excel 2016~2021 (`.FLOAT`)
* **Microsoft 365 / Excel 2024 이상:** `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)`로 정렬·필터와 연동되는 네이티브 **인셀 이미지**를 생성합니다.
* **Excel 2016 / 2019 / 2021:** 뒤에 **`.FLOAT`**(`=SMILES2IMG.FLOAT(A2)`)를 붙여 셀에 고정되는 그림 도형(`xlMoveAndSize`)으로 생성합니다. 물성·SAR 계산 수식(`=SMILES2MW`, `=SMILES.MATCH` 등)은 모든 Excel 버전에서 `.FLOAT` 없이 동일하게 작동합니다. 자세한 내용은 [구버전 Excel (.FLOAT)](/ko/excel/legacy-float)를 참고하세요.
:::

## 섹션별 바로가기

| 섹션 | 세부 페이지 | 핵심 수식 · 리본 도구 |
| :--- | :--- | :--- |
| **시작하기** | [**예제 파일 (.xlsx)**](/ko/excel/playbook) · [**활용 사례**](/ko/excel/use-cases) · [**49개 함수 전체 개요**](/ko/excel/formulas) | `=<입력>2<출력>` 통합 규격, 6가지 실전 시나리오 & 10개 탭 실습 파일 |
| **구조식 이미지** | [**SMILES · InChI**](/ko/excel/structures) · [**시약명 · 용매 · CAS**](/ko/excel/name-cas) · [**합성 반응식 (KET)**](/ko/excel/reactions) · [**스타일 · 하이라이트**](/ko/excel/styles-reactions) | `=SMILES2IMG`, `=NAME2IMG`, `=INCHI2IMG`, `=KET2IMG`, `.FLOAT` |
| **물성 · 3D 지표** | [**화학식 · 분자량 · 정밀질량**](/ko/excel/properties) · [**LogP · TPSA · 3D**](/ko/excel/lipinski-adme) | `=SMILES2MW`, `=INCHI2LOGP`, `=KET2TPSA`, `=SMILES2FSP3` |
| **SAR · 식별자** | [**부분구조 검색 (SMARTS)**](/ko/excel/sar) · [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) | `=SMILES.MATCH`, `=INCHI.MATCH`, `=KET2SMILES`, `=SMILES2INCHIKEY` |
| **시트 도구 · 3D** | [**SDF 파일 가져오기**](/ko/excel/sdf) · [**일괄 스타일 · TSV · 3D**](/ko/excel/tools-export) · [**리본 · 우클릭 메뉴**](/ko/excel/ribbon) | `[Import SDF]`, `[Batch Style]`, `[Copy TSV]`, `[3D Structure (Mol*)]` |
