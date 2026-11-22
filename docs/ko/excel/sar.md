# 부분구조 · 유사도 검색 (SAR)

특정 작용기나 골격 포함 여부를 **`TRUE` / `FALSE` 논리값 (`=SMILES.MATCH`)** 으로 판별하고, 기준 리드 물질과의 **2D 지문 유사도 점수 (`0.0~1.0` 실수, `=SMILES.SIMILARITY`)** 를 출력합니다. 동일한 쿼리 셀을 `...2IMG` 의 `"hl:"` 옵션에 연결하면 일치하는 원자만 구조식 안에서 컬러로 자동 강조됩니다.

<FormulaTable
  title="부분구조 · 유사도 · 하이라이트 시트"
  :columns='["대상 분자", "기준 구조", "수식"]'
  :rows='[
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES.MATCH(A2, \"c1ccccc1\")"], desc: "특정 골격(벤젠) 포함 여부 논리 판별", out: "TRUE" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1C(=O)O", "=SMILES.SIMILARITY(A3, $B$3)"], desc: "기준 구조($B$3) 대비 2D Tanimoto 유사도", out: "0.00~1.00" },
    { cells: ["CC(=O)Oc1ccccc1C(=O)O", "", "=SMILES2IMG(A4, \"\", \"hl:c1ccccc1\")"], desc: "일치하는 벤젠 원자·결합 컬러 강조", out: "하이라이트 2D" }
  ]'
/>

<FeatureMedia id="excel-sar-screening" />

## 1. 쿼리 셀(`$B$3`) 연동 동적 SAR 대시보드 패턴

수식마다 SMARTS 문자열을 직접 적는 대신, 상단 고정 셀(예: `$B$3`) 하나에 찾을 골격을 적어두고 참조하면 **`$B$3` 값만 바꿨을 때 시트 전체 필터·하이라이트·유사도가 동시에 갱신**됩니다:

| 열 역할 | `$B$3` 참조 수식 | `$B$3` 셀을 수정할 때의 동작 |
| :--- | :--- | :--- |
| **골격 포함 여부 (`TRUE`/`FALSE`)** | `=SMILES.MATCH(A6, $B$3)` | 새 쿼리 골격을 포함한 행만 즉시 `TRUE`로 판별 |
| **골격 컬러 하이라이트 구조식** | `=SMILES2IMG(A6, "", "hl:"&$B$3)` | 일치하는 원자·결합만 구조식 안에서 컬러로 자동 강조 |
| **2D 구조 유사도 점수 (`0.0~1.0`)** | `=SMILES.SIMILARITY(A6, $B$3, "tanimoto")` | 기준 구조와의 유사도 점수 산출 (`"tanimoto"` 기본, `"tversky"`, `"euclid-sub"` 지원) |

Lipinski 약물성 판정과 부분구조 조건을 결합하면 핵심 약리작용단(Pharmacophore)을 가진 약물성 후보만 자동 선별할 수 있습니다:

```text
=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "우선 후보", "제외")
```

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

- [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) — 시약명·CAS·SMILES·InChIKey 표준화 및 중복 화합물 제거
- [**Lipinski · ADME · 3D**](/ko/excel/lipinski-adme) — SAR 골격 필터와 Lipinski Ro5, TPSA, Fsp3 약물성 기준 결합하기
- [**SDF 파일 입출력**](/ko/excel/sdf) — 외부 `.sdf` 화합물 라이브러리를 12열 라이브 엑셀 SAR 시트로 한 번에 불러오기
- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — `hl:` 강조를 `"bw"`, `"num"`, 다크 모드와 함께 조합하기
