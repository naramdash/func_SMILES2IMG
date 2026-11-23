# LogP · TPSA · 3D 입체 지표

지용성 계수(`LogP`), 위상 극성 표면적(`TPSA`), 중원자·고리 골격 지표 및 3D 입체 탄소 비율(`Fsp3`)을 엑셀 시트에서 **인터넷 없이 0ms 순수 오프라인**으로 계산합니다. **`SMILES` · `InChI` · `KET` 세 가지 표기법 모두 동일한 수식 세트를 완벽하게 지원합니다.**

<FormulaTable
  title="LogP · TPSA · 3D 복잡도 시트"
  :columns='["SMILES", "수식"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2LOGP(A2)", desc: "Wildman-Crippen 지용성 계수 (LogP)", out: "1.31" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2TPSA(A3)", desc: "위상 극성 표면적 TPSA (Å²)", out: "63.60" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2AROMATIC(A4)", desc: "방향족 고리 개수", out: "1" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FSP3(A5)", desc: "전체 탄소 중 sp³ 입체 탄소 비율 (3D)", out: "0.11" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. LogP · TPSA · 3D 복잡도 전체 목록 (`6종` × `3개 포맷` = `18종`)

모든 지표는 **`SMILES2*`**, **`INCHI2*`**, **`KET2*`** 세 가지 입력 전용 함수로 제공되며 엑셀 기본 정수(`Integer`)·실수(`Double`)를 반환합니다 (`SMILES2*`는 InChI 및 단일 분자 KET도 자동 감지하며, `INCHI2*`·`KET2*`는 각각 해당 포맷이 아니면 `#VALUE!`로 정확히 차단합니다):

| 지표 분류 | SMILES 수식 | InChI / KET 전용 수식 | 반환 타입 | 아스피린 (`A2`) | 참고 기준 |
| :--- | :--- | :--- | :---: | :---: | :--- |
| **지용성 (LogP)** | `=SMILES2LOGP(A2)` | `=INCHI2LOGP(A2)` · `=KET2LOGP(A2)` | `Double` | `1.31` | Wildman-Crippen 옥탄올-물 분배계수 (`≤ 5.0`) |
| **극성 표면적 (TPSA)** | `=SMILES2TPSA(A2, [S/P포함])` | `=INCHI2TPSA(A2, [S/P])` · `=KET2TPSA(A2, [S/P])` | `Double` | `63.60` | `≤ 140 Å²` 경구; `≤ 90 Å²` CNS BBB (`TRUE`는 S/P 원자 포함) |
| **중원자 수 (Heavy Atoms)** | `=SMILES2HEAVY(A2)` | `=INCHI2HEAVY(A2)` · `=KET2HEAVY(A2)` | `Integer` | `13` | 수소를 제외한 비수소 원자 개수 (`countHeavyAtoms`) |
| **전체 고리 수 (SSSR)** | `=SMILES2RINGS(A2)` | `=INCHI2RINGS(A2)` · `=KET2RINGS(A2)` | `Integer` | `1` | 분자 내 최소 고리 집합 전체 개수 |
| **방향족 고리 수** | `=SMILES2AROMATIC(A2)` | `=INCHI2AROMATIC(A2)` · `=KET2AROMATIC(A2)` | `Integer` | `1` | **≤ 3** 권장 *(EPAM Indigo generic 모델 기준)* |
| **입체 탄소 비율 (Fsp³)** | `=SMILES2FSP3(A2)` | `=INCHI2FSP3(A2)` · `=KET2FSP3(A2)` | `Double` | `0.11` | **≥ 0.25** ("Escape from Flatland" `sp³ C / 전체 C`; 무탄소 물질은 `0.00`) |

::: info 화학적 정확성 및 계산 규격 안내 (EPAM Indigo 엔진 기준)
* **TPSA 모드 및 허용 인자:** 기본값 `=SMILES2TPSA(A2)` 는 Ertl 표준(N, O 원자만 포함, RDKit 기본값과 동일)을 따르며, 2번째 인수로 `TRUE`(또는 `1`, `"yes"`, `"sp"`, `"s"`, `"p"`, `"pubchem"`, `"all"`)를 주면 S, P 원자까지 포함해 PubChem 기준값과 맞출 수 있습니다 (`=SMILES2TPSA(A2, TRUE)`).
* **전하 표기 · 호변이성질체 · 염 주의사항:** `LogP`와 `TPSA`는 양성자화/형식 전하 표기(예: 양성이온 vs 중성 아미노산, `[N+](=O)[O-]` vs `N(=O)=O`, `[S+]([O-])` vs `S=O`)와 표준 `InChI=1S` 이동 수소(mobile-H)의 아마이드/에놀 호변이성질체 재구성 여부에 민감합니다(정확한 호변이성질체 기준이 필요할 때는 `SMILES` 또는 `KET` 권장). 또한 염·짝이온이나 유기금속·무기물(`NaCl`, `Cisplatin` 등)은 그려진 전체 단편에 대해 Indigo Crippen fallback 계수가 계산되므로 실험 분배계수로 직접 해석하지 마세요.
* **모호한 HBD/HBA/Lipinski/카이랄 래퍼 미제공:** 툴킷마다 정의가 달라 불일치가 잦은 HBD/HBA/회전결합·블랙박스 `"PASS"`·명시 표기만 세는 카이랄 개수 대신, 분자량(`SMILES2MW`), `LogP`, `TPSA`, `Fsp3`, `Aromatic` 수치를 엑셀 `IF(AND(...))` 수식에 투명하게 조합해 사용하도록 설계했습니다.
* **쿼리·반응식 가드:** R-기(`*`) 또는 반응식(`>>`)을 스칼라 지표 함수에 넣으면 모두 `#N/A`를 반환합니다.
:::

## 2. 투명하고 재현 가능한 Hit-to-Lead 선별 수식 (`IF` / `AND`)

반환값이 표준 엑셀 숫자이므로 `IF`, `AND`, `FILTER` 와 결합해 자체 기준대로 후보를 투명하게 선별할 수 있습니다:

| 신약 선별 규칙 | 복사해서 쓰는 엑셀 수식 (`A2` = SMILES/InChI/KET) |
| :--- | :--- |
| **경구 물성 윈도우 (`MW ≤ 500`, `LogP ≤ 5`, `TPSA ≤ 140`)** | `=IF(AND(SMILES2MW(A2)<=500, SMILES2LOGP(A2)<=5, SMILES2TPSA(A2)<=140), "PASS", "FAIL")` |
| **CNS BBB 투과 후보 (`TPSA ≤ 90`, `MW ≤ 450`, `LogP ≤ 4`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450, SMILES2LOGP(A2)<=4), "CNS 후보", "일반")` |
| **3D 입체성 우수 리드 (`Fsp3 ≥ 0.25`, `방향족 고리 ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "입체 우수", "평면 과다")` |
| **물성 + 핵심 골격 동시 선별** | `=IF(AND(SMILES2MW(A2)<=500, SMILES2TPSA(A2)<=120, SMILES.MATCH(A2, "c1ccncc1")), "우선 후보", "제외")` |

## 관련 가이드

- [**화학식 · 분자량 · 정밀질량**](/ko/excel/properties) — 분자량 기반 시약 칭량(`mg`) 및 HRMS *m/z* 계산표
- [**부분구조 검색 (SMARTS)**](/ko/excel/sar) — ADME 선별식에 `=SMILES.MATCH` · `=INCHI.MATCH` · `=KET.MATCH` 결합
- [**예제 파일 (.xlsx)**](/ko/excel/playbook) — `2. Properties & 3D` 탭에서 완성된 스코어링 대시보드 실습
