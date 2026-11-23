# 부분구조 검색 (SAR & SMARTS)

**`=SMILES.MATCH`**, **`=INCHI.MATCH`**, **`=KET.MATCH`** 수식으로 3가지 표기법 모두에서 특정 작용기·골격 포함 여부를 **`TRUE` / `FALSE` 논리값**으로 판별합니다. 같은 쿼리 셀을 `...2IMG` 의 `"hl:"` 옵션에 연결하면 일치하는 원자·결합만 구조식 안에서 컬러로 자동 강조됩니다.

<FormulaTable
  title="부분구조 · 하이라이트 SAR 시트"
  :columns='["대상 분자", "쿼리", "수식"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1", "=SMILES.MATCH(A2, $B$2)"], desc: "특정 골격(벤젠) 포함 여부 논리 판별", out: "TRUE" },
    { cells: ["InChI=1S/C9H8O4/c1-6(10)13-8-5-3-2-4-7(8)9(11)12/h2-5H,1H3,(H,11,12)", "", "=INCHI.MATCH(A3, $B$2)"], desc: "InChI 셀 직접 부분구조 검색", out: "TRUE" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES2IMG(A4, \"\", \"hl:\"&$B$2)"], desc: "일치하는 벤젠 원자·결합 컬러 강조", out: "하이라이트 2D" }
  ]'
/>

<FeatureMedia id="excel-sar-screening" />

## 1. 쿼리 셀(`$B$3`) 연동 동적 SAR 대시보드 패턴

수식마다 SMARTS 문자열을 직접 적는 대신 상단 고정 셀(예: `$B$3`) 하나에 찾을 골격을 적어 두고 참조하면, **`$B$3` 값만 바꿨을 때 시트 전체 필터와 구조식 하이라이트가 동시에 갱신**됩니다:

| 열 역할 | `$B$3` 참조 수식 | `$B$3` 셀을 수정할 때의 동작 |
| :--- | :--- | :--- |
| **SMILES 부분구조 (`TRUE`/`FALSE`)** | `=SMILES.MATCH(A6, $B$3)` | 새 쿼리 골격을 포함한 SMILES 행만 즉시 `TRUE` 로 판별 |
| **InChI / KET 부분구조 검색** | `=INCHI.MATCH(A6, $B$3)` · `=KET.MATCH(A6, $B$3)` | InChI 또는 KET 캔버스 열을 직접 SMARTS 검색 |
| **골격 컬러 하이라이트 구조식** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | 일치하는 원자·결합만 구조식 안에서 컬러로 자동 강조 |

물성 조건과 부분구조 조건을 결합하면 핵심 약리작용단(Pharmacophore)을 가진 우수 후보만 투명하게 선별할 수 있습니다:

```text
=IF(AND(SMILES2MW(A2)<=500, SMILES2LOGP(A2)<=4, SMILES.MATCH(A2, "c1ccncc1")), "우선 후보", "제외")
```

::: info SMARTS 문법 규격 안내
`*.MATCH` 와 `hl:` 는 표준 Daylight/RDKit **SMARTS** 를 우선 해석합니다. 대문자 `C`, `N` 은 지방족(aliphatic) 원자를, 소문자 `c`, `n` 은 방향족 원자를 의미하므로 둘 다 찾으려면 원자 번호 표기(`[#6]`, `[#7]`)를 사용하세요. 대상 분자는 Indigo `generic` 방향족 모델로 정규화되므로 케쿨레 표기와 방향족 SMILES가 동일하게 매칭됩니다.
:::

## 2. 바로 복사해 쓰는 12대 필수 작용기 SMARTS 패턴 사전

실무에서 가장 자주 검색하는 12가지 작용기·고리의 포함 판정(`=SMILES.MATCH`) 및 하이라이트(`hl:`) 수식입니다:

| 작용기 / 고리 골격 | SMARTS 패턴 | 포함 여부 (`TRUE` / `FALSE`) | 하이라이트 구조식 수식 |
| :--- | :--- | :--- | :--- |
| **벤젠 고리 (방향족)** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **피리딘 고리 (헤테로방향족)** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **카복실산 (`-COOH`)** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **에스터 (`-COO-`)** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **아마이드 (`-CONH-`)** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **설폰아마이드 (`-SO₂NH-`)** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **케톤 (`-C(=O)-`)** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **알데하이드 (`-CHO`)** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **페놀 (방향족 `-OH`)** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **나이트로기 (`-NO₂`)** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **할로겐 원소 (`F, Cl, Br, I`)** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **트라이플루오로메틸 (`-CF₃`)** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

## 관련 가이드

- [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) — 시약명·CAS·SMILES·KET·InChIKey 표준화 및 중복 화합물 제거
- [**LogP · TPSA · 3D**](/ko/excel/lipinski-adme) — SAR 골격 필터와 LogP, TPSA, Fsp3 지표 결합하기
- [**SDF 파일 가져오기**](/ko/excel/sdf) — 외부 `.sdf` 화합물 라이브러리를 11열 라이브 엑셀 SAR 시트로 한 번에 불러오기
- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — `hl:` 강조를 `"bw"`, `"num"`, 다크 모드와 함께 조합하기
