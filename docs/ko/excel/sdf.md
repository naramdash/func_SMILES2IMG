# SDF 파일 가져오기 · 내보내기

다중 화합물 `.sdf` / `.mol` 화학 데이터베이스를 12열 라이브 수식 시트로 한 번에 불러오거나, 작업한 시트를 표준 V2000 `.sdf` 파일로 내보냅니다.

<FeatureMedia id="excel-sdf-workflow" />

::: tip 샘플 `.sdf` 데이터베이스로 바로 테스트하기
<a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf"><strong>sample_compounds.sdf</strong></a>를 내려받아 Excel 상단 **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** 메뉴로 열어보세요.
:::

## 1. `.sdf` 파일 가져오기 (Import)

1. 상단 리본에서 **`[SMILES]` ➔ `[SDF File]` ➔ `[Import SDF File (.sdf)...]`** 메뉴를 클릭합니다.
2. `.sdf` 또는 `.mol` 파일을 선택합니다.
3. 행 높이·열 너비가 자동 맞춤된 새 시트가 생성되며 **12개 라이브 수식 열**이 자동으로 채워집니다:
   - `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski`

## 2. 엑셀 시트를 `.sdf`로 내보내기 (Export)

1. 구조식 또는 SMILES 열이 있는 시트를 엽니다.
2. **`[SMILES]` ➔ `[SDF File]` ➔ `[Export Sheet to SDF (.sdf)...]`** 메뉴를 클릭합니다.
3. 저장 경로를 지정하면 RDKit, KNIME, DataWarrior, ChemDraw와 호환되는 표준 **V2000 `.sdf`** 파일이 생성됩니다.
