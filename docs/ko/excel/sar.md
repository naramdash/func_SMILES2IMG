# 부분구조 · 유사도 (SAR)

특정 작용기나 골격을 포함한 화합물을 필터링하고(`=SMILES.MATCH`), 구조식 안에서 해당 골격만 컬러로 강조하며(`hl:`), 기준 물질과의 2D 유사도를 계산합니다(`=SMILES.SIMILARITY`).

```text
=SMILES.MATCH(A2, "c1ccccc1")
=SMILES.SIMILARITY(A2, $B$2)
=SMILES2IMG(A2, "", "hl:c1ccccc1")
```

<FeatureMedia id="excel-sar-screening" />

## 자주 쓰는 12대 작용기 SMARTS 패턴 사전

| 작용기 / 고리 | SMARTS 패턴 | 포함 여부 (`TRUE` / `FALSE`) | 하이라이트 구조식 수식 |
| :--- | :--- | :--- | :--- |
| **벤젠 고리** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **피리딘 고리** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **카복실산** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **에스터** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **아마이드** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **설폰아마이드** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **케톤** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **알데하이드** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **페놀** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **나이트로기** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **할로겐** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **트라이플루오로메틸** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

## 유사도 지표 및 `IF` 조건 결합

* `=SMILES.SIMILARITY(mol1, mol2, [metric])`은 `"tanimoto"`(기본값), `"tversky"`, `"euclid-sub"` 지표를 지원하며 `0.0`~`1.0` 숫자를 반환합니다.
* Lipinski 판정과 부분구조 조건을 결합해 우선순위 후보를 선별할 수 있습니다:

```text
=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "우선 후보", "제외")
```
