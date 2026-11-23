# Excel 활용 사례

**`func_SMILES2IMG`** 로 Microsoft Excel을 연구실 실무 워크벤치로 활용하는 6가지 대표 시나리오입니다. 아래 표는 **`전체 복사`** 후 엑셀 `A1` 셀에 붙여넣으면 수식과 입력값이 바로 동작합니다.

| 시나리오 | 해결하는 작업 | 핵심 기능 |
| :--- | :--- | :--- |
| [**1. 유기 합성 ELN · 당량 칭량**](#synthesis-eln) | 시약명·CAS만으로 SMILES, 구조식, 분자량, 칭량 `mg` 자동 계산 | `=NAME2SMILES`, `=SMILES2MW` |
| [**2. Hit-to-Lead SAR 선별**](#sar-triage) | 기준 쿼리 셀(`$B$2`)로 부분구조 필터·하이라이트·LogP 동시 선별 | `=SMILES.MATCH`, `"hl:"` |
| [**3. 시약 재고 · 중복 화합물 제거**](#deduplication) | 표기 순서가 다른 SMILES를 `InChIKey` 로 표준화해 중복 검출·병합 | `=SMILES2INCHIKEY` |
| [**4. SDF 불러오기 · 보고서 연동**](#sdf-handoff) | `.sdf` 파일을 11열 수식 시트로 열고 Word 논문·PPT 발표로 바로 이동 | `[Import SDF]`, `[Batch Style]` |
| [**5. R-그룹 유도체 자동 조합**](#rgroup-enumeration) | `SUBSTITUTE` 치환기 조합만으로 구조식·MW·화학식 오프라인 산출 | `=SUBSTITUTE`, `=SMILES2IMG` |
| [**6. 드롭다운 리드 비교 · FILTER 보드**](#dynamic-comparison) | `XLOOKUP`·`FILTER` 로 인셀 구조식 이미지를 동적으로 추출·비교 (365+) | `=XLOOKUP`, `=FILTER` |

---

## 1. 유기 합성 ELN · 당량 칭량 {#synthesis-eln}

반응 셋업 때마다 ChemDraw·PubChem·계산기를 오갈 필요 없이, 상단에 **`[SMILES]` ➔ `[Insert / Edit (KET)]`** 로 반응식을 붙여 두고 `A` 열에 시약명이나 CAS 번호만 입력해 SMILES, 구조식, 분자량(`MW`), 칭량값(`mg`)을 자동 연동합니다.

<FeatureMedia id="excel-reactions-ket" />

1. 아래 표를 **`전체 복사`** 해 엑셀 **`A1`** 셀에 붙여넣으세요 (Excel 2016~2021은 `=SMILES2IMG` 뒤에 **`.FLOAT`** 추가).
2. 한계 시약 `E2` 셀(`5.0` mmol) 숫자만 바꾸면 전 시약의 필요 칭량값(`F2:F4`)이 즉시 재계산됩니다.
3. 시트 상단 셀에는 **`[Insert / Edit (KET)]`** 로 더블클릭 편집 가능한 반응식을 배치하세요.

<FormulaTable
  title="ELN 당량 칭량 계산기 시트"
  :columns='["시약명 / CAS", "SMILES", "구조식", "분자량(MW)", "mmol", "칭량(mg)"]'
  :rows='[
    { cells: ["Salicylic Acid", "=NAME2SMILES(A2)", "=SMILES2IMG(B2)", "=SMILES2MW(B2)", "5.0", "=ROUND(D2*E2, 1)"], desc: "한계 시약 (1.0 eq)", out: "690.6 mg" },
    { cells: ["Acetic Anhydride", "=NAME2SMILES(A3)", "=SMILES2IMG(B3)", "=SMILES2MW(B3)", "=E2*1.5", "=ROUND(D3*E3, 1)"], desc: "과량 시약 (1.5 eq)", out: "765.7 mg" },
    { cells: ["50-78-2", "=NAME2SMILES(A4)", "=SMILES2IMG(B4)", "=SMILES2MW(B4)", "=E2", "=ROUND(D4*E4, 1)"], desc: "이론 아스피린 생성량", out: "900.8 mg" }
  ]'
/>

* **관련 가이드:** [**화학식 · 분자량 · 정밀질량**](/ko/excel/properties) · [**합성 반응식 (KET)**](/ko/excel/reactions) · Playbook **Tab 6 & Tab 7**

---

## 2. Hit-to-Lead SAR 선별 {#sar-triage}

수백 행의 유도체 시트를 고정 쿼리 셀(`$B$2`) 하나에 연결하면, 찾을 SMARTS 패턴을 바꿨을 때 부분구조 통과 여부, 골격 컬러 강조, LogP 기준 선별 등급이 동시에 갱신됩니다.

<FeatureMedia id="excel-sar-screening" />

1. 아래 표를 **`A1`** 셀에 붙여넣습니다 (`$B$2` 셀이 동적 쿼리 SMARTS `c1ccncc1` 역할).
2. `$B$2` 값을 `c1ccccc1` 등으로 수정하면 `C:E` 열의 일치 판정·하이라이트·우선순위가 일괄 업데이트됩니다.
3. Microsoft 365 / Excel 2024+에서는 표를 정렬·필터링해도 인셀 구조식이 행을 정확히 따라옵니다.

<FormulaTable
  title="동적 쿼리 SAR 선별 시트"
  :columns='["대상 SMILES", "쿼리 ($B$2)", "포함 여부", "하이라이트 2D", "선별 등급"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccncc1", "=SMILES.MATCH(A2, $B$2)", "=SMILES2IMG(A2, \"\", \"hl:\"&$B$2)", "=IF(AND(SMILES2LOGP(A2)<=4.0, SMILES.MATCH(A2, $B$2)), \"Priority\", \"Monitor\")"], desc: "아스피린 vs. 피리딘 쿼리", out: "Monitor" },
    { cells: ["CN1CCCC1c2cccnc2", "", "=SMILES.MATCH(A3, $B$2)", "=SMILES2IMG(A3, \"\", \"hl:\"&$B$2)", "=IF(AND(SMILES2LOGP(A3)<=4.0, SMILES.MATCH(A3, $B$2)), \"Priority\", \"Monitor\")"], desc: "니코틴 — $B$2 피리딘 포함", out: "Priority" }
  ]'
/>

* **관련 가이드:** [**부분구조 검색 (SAR)**](/ko/excel/sar) · [**LogP · TPSA · 3D 복잡도**](/ko/excel/lipinski-adme) · Playbook **Tab 2 & Tab 3**

---

## 3. 시약 재고 · 중복 화합물 제거 {#deduplication}

공급처나 논문마다 동일 분자의 SMILES 원자 순서가 달라(`c1ccccc1C(=O)O` vs `O=C(O)c1ccccc1`) 일반 텍스트 비교로는 중복을 놓치기 쉽습니다. 원자 순서와 무관한 27자리 **`InChIKey`** 로 정규화하면 `COUNTIF` 중복 표시, `XLOOKUP` 시트 병합, **중복된 항목 제거**를 정확하게 수행할 수 있습니다.

<FeatureMedia id="excel-conversions" />

1. 아래 표를 **`A1`** 셀에 붙여넣으면 표기 순서가 다른 두 벤조산과 아스피린을 비교합니다.
2. `COUNTIF($C$2:$C$4, C2)>1` 수식을 전체 행으로 채워 화학적으로 동일한 중복 항목을 즉시 검출합니다.
3. 다른 공급처 시트 단가·활성 병합 시 `SMILES` 대신 `InChIKey` 열을 `XLOOKUP` 기준 키로 사용하세요.

<FormulaTable
  title="InChIKey 중복 제거 시트"
  :columns='["입력 SMILES / 시약명", "표준 SMILES", "InChIKey", "중복 판별"]'
  :rows='[
    { cells: ["c1ccccc1C(=O)O", "=INCHI2SMILES(SMILES2INCHI(A2))", "=SMILES2INCHIKEY(A2)", "=IF(COUNTIF($C$2:$C$4, C2)>1, \"중복\", \"고유\")"], desc: "벤조산 (고리 우선 SMILES)", out: "중복" },
    { cells: ["O=C(O)c1ccccc1", "=INCHI2SMILES(SMILES2INCHI(A3))", "=SMILES2INCHIKEY(A3)", "=IF(COUNTIF($C$2:$C$4, C3)>1, \"중복\", \"고유\")"], desc: "동일 분자 (카복실산 우선)", out: "중복" },
    { cells: ["Aspirin", "=NAME2SMILES(A4)", "=SMILES2INCHIKEY(B4)", "=IF(COUNTIF($C$2:$C$4, C4)>1, \"중복\", \"고유\")"], desc: "시약명 변환 아스피린", out: "고유" }
  ]'
/>

* **관련 가이드:** [**시약명 · 용매 · CAS**](/ko/excel/name-cas) · [**식별자 변환 · InChIKey**](/ko/excel/data-sdf) · Playbook **Tab 4**

---

## 4. SDF 불러오기 · 보고서 연동 {#sdf-handoff}

RDKit, KNIME, DataWarrior, PubChem에서 내보낸 `.sdf` 화학 DB를 엑셀 수식 표로 한 번에 열고, 선별한 후보만 Word 논문·PowerPoint 다크 모드 슬라이드로 손실 없이 넘길 수 있습니다.

<FeatureMedia id="excel-sdf-workflow" />

1. **11열 수식 표 생성:** **`[SMILES]` ➔ `[Import SDF]`** 를 누르면 `Name`, `Structure`, `SMILES`, `Formula`, `MW (g/mol)`, `Exact Mass`, `LogP`, `TPSA`, `Rings`, `Aromatic Rings`, `Fsp3` 열이 자동 생성됩니다.
2. **스타일 일괄 통일:** 상위 히트만 필터링한 뒤 **`[SMILES]` ➔ `[Batch Style]` ➔ `[Publication B&W]`** (또는 `Dark Mode`)를 적용합니다.
3. **Word · PowerPoint 연동:**
   - **Word 논문 표로:** 물성 행을 복사해 붙이거나, 구조식 셀에서 **`[Copy SMILES]`** / **`[Copy KET]`** 후 Word에서 **`[Paste Structure]`** 를 누르세요 ([Word 활용 사례](/ko/word/use-cases)).
   - **PowerPoint 발표로:** 셀에서 복사 후 슬라이드에 **`[Paste Structure]`** ➔ **`[Batch Style]` ➔ `[Dark Mode]`** 를 적용하세요 ([PowerPoint 활용 사례](/ko/powerpoint/use-cases)).

* **관련 가이드:** [**SDF 파일 가져오기**](/ko/excel/sdf) · [**일괄 스타일 · TSV · 3D**](/ko/excel/tools-export) · Playbook **Tab 9**

---

## 5. R-그룹 유도체 자동 조합 {#rgroup-enumeration}

Python이나 RDKit 없이 엑셀 기본 함수 **`SUBSTITUTE`** 만으로 치환기별 구조식과 물성을 시뮬레이션할 수 있습니다. `A` 열에 치환기 SMILES 조각만 추가하면 유도체 구조식·분자량·화학식·LogP가 로컬에서 계산됩니다.

<FeatureMedia id="excel-cheminformatics" />

1. 아래 유도체 조합 시트를 **`A1`** 셀에 붙여넣습니다 (`$B$2` 셀이 `[R]` 골격 템플릿).
2. `A` 열 아래로 `Cl`, `OC`, `C#N` 등 다른 치환기 조각을 추가해 `C:F` 열 수식을 드래그해 채우세요.
3. `=SMILES2LOGP(C2)` · `=SMILES2TPSA(C2)` 열을 추가하면 합성 전 물성 균형을 바로 비교할 수 있습니다.

<FormulaTable
  title="R-그룹 유도체 자동 조합 시트"
  :columns='["R-그룹", "템플릿 ($B$2)", "조합 SMILES", "2D 구조식", "MW", "화학식"]'
  :rows='[
    { cells: ["O", "CC(=O)Nc1ccc([R])cc1", "=SUBSTITUTE($B$2, \"[R]\", A2)", "=SMILES2IMG(C2)", "=SMILES2MW(C2)", "=SMILES2FORMULA(C2)"], desc: "4-OH 유도체 (파라세타몰)", out: "151.16 · C8H9NO2" },
    { cells: ["F", "", "=SUBSTITUTE($B$2, \"[R]\", A3)", "=SMILES2IMG(C3)", "=SMILES2MW(C3)", "=SMILES2FORMULA(C3)"], desc: "4-Fluoro 유도체", out: "153.16 · C8H8FNO" },
    { cells: ["C(F)(F)F", "", "=SUBSTITUTE($B$2, \"[R]\", A4)", "=SMILES2IMG(C4)", "=SMILES2MW(C4)", "=SMILES2FORMULA(C4)"], desc: "4-CF3 유도체", out: "203.16 · C9H8F3NO" }
  ]'
/>

* **관련 가이드:** [**SMILES · InChI**](/ko/excel/structures) · [**LogP · TPSA · 3D 복잡도**](/ko/excel/lipinski-adme)

---

## 6. 드롭다운 리드 비교 · FILTER 보드 {#dynamic-comparison}

**Microsoft 365 / Excel 2024+** 에서 `=SMILES2IMG` 는 셀 값 자체인 네이티브 **인셀 이미지(`CellImage`)** 를 반환하므로, 엑셀의 `XLOOKUP` · `FILTER` 수식이 구조식 그림까지 동적으로 끌어올 수 있습니다.

<FeatureMedia id="excel-formula-table" />

1. **드롭다운 1:1 리드 비교 카드:** `B2`와 `C2` 셀에 `Library` 시트 화합물명 드롭다운을 두고, `=XLOOKUP(B2, Library!$A:$A, Library!$C:$C)` 로 라이브 구조식 셀을 불러온 뒤 `=SMILES.MATCH(B3, Library!$B$1)` 및 `ΔMW` / `ΔLogP` 비교 행을 연결합니다.
2. **실시간 조건별 `FILTER` 보드:** 조건을 만족하는 후보만 구조식 그림과 함께 자동 추출할 수 있습니다:
   ```text
   =FILTER(Library!A2:F200, (Library!E2:E200<=4.0)*(Library!F2:F200<=90))
   ```

::: info Microsoft 365 / Excel 2024+ 전용 기능
`XLOOKUP` · `FILTER` 로 구조식 그림까지 이동시키는 기능은 네이티브 인셀 이미지(`=SMILES2IMG`) 환경에서 동작하며, 구버전 `.FLOAT` 도형에는 적용되지 않습니다.
:::

* **관련 가이드:** [**SMILES · InChI**](/ko/excel/structures) · [**부분구조 검색 (SAR)**](/ko/excel/sar)

---

## 다음 단계

- [**10개 탭 Excel Playbook (.xlsx) 실습하기**](/ko/excel/playbook) — 위 템플릿이 모두 담긴 통합 실습 시트를 바로 열어 보세요.
- [**설치 가이드**](/ko/guide/install) — Excel·Word·PowerPoint 통합 설치 명령어를 확인하세요.
