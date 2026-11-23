# SDF 파일 가져오기 (`.sdf` / `.mol`)

RDKit, KNIME, DataWarrior, ChemDraw, PubChem 등에서 내보낸 **다중 화합물 `.sdf` / `.mol` 화학 라이브러리**를 클릭 한 번으로 **11열 라이브 수식 시트**로 불러옵니다.

<FeatureMedia id="excel-sdf-workflow" />

::: tip 샘플 `.sdf` 데이터베이스로 바로 테스트하기
<a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf"><strong>sample_compounds.sdf</strong></a> 를 내려받아 Excel 상단 **`[SMILES]` ➔ `[Import SDF]`** 메뉴로 바로 열어보세요.
:::

## 1. `.sdf` / `.mol` 파일 가져오기 (11열 라이브 시트 자동 생성)

1. 상단 리본에서 **`[SMILES]` ➔ `[Import SDF]`** 를 클릭합니다.
2. `.sdf` 또는 `.mol` 파일을 선택합니다 (최대 500개 화합물 자동 가져오기).
3. 행 높이·열 너비가 자동 맞춤된 시트 표가 생성되며 아래 **11개 라이브 열**이 즉시 채워집니다:
   - `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `TPSA`, `Rings`, `Aromatic Rings`, `Fsp3`

불러온 모든 행은 표준 엑셀 수식(`=SMILES2IMG`, `=SMILES2MW`, `=SMILES2LOGP`, `=SMILES2TPSA` 등)으로 채워지므로 더블클릭 편집과 정렬·필터가 즉시 작동합니다.

## 관련 가이드

- [**부분구조 검색 (SMARTS)**](/ko/excel/sar) — 불러온 `.sdf` 시트에서 `=SMILES.MATCH` 와 `hl:` 로 후보 물질 선별
- [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) — 여러 `.sdf` DB 를 합친 뒤 `=SMILES2INCHIKEY` 로 중복 제거하기
- [**일괄 스타일 · TSV · 3D**](/ko/excel/tools-export) — 가져온 시트 전체를 원클릭으로 ACS 흑백이나 다크 모드로 손실 없이 통일하기
