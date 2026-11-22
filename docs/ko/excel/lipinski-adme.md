# Lipinski · ADME · 3D 입체 지표

경구 흡수 가능성(Lipinski Rule of 5), 뇌혈관장벽 투과도(`TPSA`), 고리 골격 지표 및 3D 입체 복잡도(`Fsp3`, 카이랄 중심 수)를 엑셀 시트에서 **인터넷 없이 0ms 순수 오프라인**으로 계산합니다.

<FormulaTable
  title="Lipinski · ADME · 3D 복잡도 시트"
  :columns='["SMILES", "수식"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LIPINSKI(A2)", desc: "Lipinski Rule of 5 약물성 종합 판정", out: "PASS" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LOGP(A3)", desc: "Wildman-Crippen 지용성 계수 (LogP)", out: "1.31" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2TPSA(A4)", desc: "위상 극성 표면적 TPSA (Å²)", out: "63.60" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FSP3(A5)", desc: "전체 탄소 중 sp³ 입체 탄소 비율 (3D)", out: "0.11" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. ADME · Lipinski · 3D 복잡도 지표 전체 목록 (`11종`)

11개 지표 함수 모두 SMILES, InChI, 단일 분자 KET 셀을 자동 인식하며 엑셀 기본 정수(`Integer`), 실수(`Double`), 또는 `"PASS"`/`"FAIL"` 문자열을 반환합니다:

| 지표 분류 | 엑셀 수식 | 반환 타입 | 아스피린 (`A2`) | 의약화학 권장 기준 |
| :--- | :--- | :---: | :---: | :--- |
| **Lipinski 종합 판정** | `=SMILES2LIPINSKI(A2)` | `String` | `PASS` | `MW ≤ 500`, `LogP ≤ 5.0`, `HBD ≤ 5`, `HBA ≤ 10` 모두 통과 시 `"PASS"` |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | `Double` | `1.31` | **≤ 5.0** (Wildman-Crippen 옥탄올-물 분배계수) |
| **극성 표면적 (TPSA)** | `=SMILES2TPSA(A2)` | `Double` | `63.60` | **≤ 140 Å²** 경구 흡수; **≤ 90 Å²** 뇌혈관장벽(CNS BBB) 투과 |
| **수소 결합 공여체 (HBD)** | `=SMILES2HBD(A2)` | `Integer` | `1` | **≤ 5** (`-OH`, `-NH` 개수) |
| **수소 결합 수용체 (HBA)** | `=SMILES2HBA(A2)` | `Integer` | `4` | **≤ 10** (`N`, `O` 원자 개수) |
| **회전 가능한 결합 (RotB)** | `=SMILES2ROTBONDS(A2)` | `Integer` | `3` | **≤ 10** (Veber 경구 생체이용률 기준) |
| **중원자 수 (Heavy Atoms)** | `=SMILES2HEAVY(A2)` | `Integer` | `13` | 수소를 제외한 비수소 원자 개수 (`countHeavyAtoms`) |
| **전체 고리 수 (SSSR)** | `=SMILES2RINGS(A2)` | `Integer` | `1` | 분자 내 최소 고리 집합 전체 개수 |
| **방향족 고리 수** | `=SMILES2AROMATIC(A2)` | `Integer` | `1` | **≤ 3** 권장 (초과 시 용해도·선택성 저하 위험) |
| **입체 탄소 비율 (Fsp³)** | `=SMILES2FSP3(A2)` | `Double` | `0.11` | **≥ 0.25** ("Escape from Flatland" 3D 입체성 지표, `0.00~1.00`) |
| **카이랄 중심 (Chiral)** | `=SMILES2CHIRAL(A2)` | `Integer` | `0` | 정의된 사면체 입체 중심(`@` / `@@`) 개수 |

## 2. 복사해서 바로 쓰는 Hit-to-Lead 선별 수식 (`IF` / `AND`)

반환값이 표준 엑셀 숫자·문자열이므로 `IF`, `AND`, `FILTER` 와 결합해 수백 행에서 우수 후보를 자동 선별할 수 있습니다:

| 신약 선별 규칙 | 복사해서 쓰는 엑셀 수식 (`A2` = SMILES/InChI/KET) |
| :--- | :--- |
| **Veber 경구 흡수성 판별 (`TPSA ≤ 140`, `RotB ≤ 10`)** | `=IF(AND(SMILES2TPSA(A2)<=140, SMILES2ROTBONDS(A2)<=10), "Veber 통과", "흡수 불량")` |
| **CNS BBB 투과 후보 (`TPSA ≤ 90`, `MW ≤ 450`, `Lipinski = PASS`)** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "CNS 후보", "일반")` |
| **3D 입체성 우수 리드 (`Fsp3 ≥ 0.25`, `방향족 고리 ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "입체 우수", "평면 과다")` |
| **약물성 + 핵심 골격 동시 선별** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "우선 후보", "제외")` |

## 관련 가이드

- [**화학식 · 분자량 · 정밀질량**](/ko/excel/properties) — 분자량 기반 시약 칭량(`mg`) 및 HRMS *m/z* 계산표
- [**부분구조 · 유사도 검색**](/ko/excel/sar) — ADME 선별식에 `=SMILES.MATCH` 골격 필터 및 Tanimoto 순위 결합
- [**예제 파일 (.xlsx)**](/ko/excel/playbook) — `2. Properties & Lipinski` 탭에서 완성된 스코어링 대시보드 실습
