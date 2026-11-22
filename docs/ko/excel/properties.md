# 화학식 · 분자량 · 정밀질량

표준 Hill 화학식, 평균 분자량(`MW`), HRMS 단일동위원소 정밀질량, 물질 이름을 **0ms 로컬 오프라인**으로 계산하고, 반환된 실제 숫자를 유기 합성 ELN 칭량 당량 계산기와 LC-MS / HRMS 이온 피크 계산표에 바로 연결합니다.

<FormulaTable
  title="화학식 · 분자량 · 정밀질량 시트 (0ms 오프라인)"
  :columns='["SMILES", "수식"]'
  :rows='[
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2FORMULA(A2)", desc: "표준 Hill 표기법 화학식 (0ms 오프라인)", out: "C9H8O4" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2MW(A3)", desc: "당량 계산용 평균 분자량 (g/mol 실수)", out: "180.16" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2EXACTMASS(A4)", desc: "HRMS 단일동위원소 정밀질량 (소수 4자리)", out: "180.0423" },
    { input: "CC(=O)Oc1ccccc1C(=O)O", formula: "=SMILES2NAME(A5)", desc: "IUPAC / PubChem 영문 물질명 (캐시 지원)", out: "Aspirin" }
  ]'
/>

<FeatureMedia id="excel-cheminformatics" />

## 1. 화학식 · 분자량 · 정밀질량 · 물질명 함수 (`12종`)

모든 함수는 단일 인수 `(셀)` 문법으로 동작하며, SMILES · InChI · 단일 분자 KET 입력을 모두 지원합니다(`INCHI2*`, `KET2*` 도 내부적으로 동일 파서를 호출합니다):

| 출력 항목 | SMILES 입력 | InChI / KET 입력 시 | 아스피린 (`A2`) | 처리 방식 및 반환 타입 |
| :--- | :--- | :--- | :---: | :--- |
| **화학식 (Hill Formula)** | `=SMILES2FORMULA(A2)` | `=INCHI2FORMULA(A2)` · `=KET2FORMULA(A2)` | `C9H8O4` | **0ms 오프라인** · 표준 힐 표기법 (`String`) |
| **분자량 (MW)** | `=SMILES2MW(A2)` | `=INCHI2MW(A2)` · `=KET2MW(A2)` | `180.16` | **0ms 오프라인** · 사칙연산 가능한 실수 `Double` (g/mol) |
| **정밀 질량 (Exact Mass)** | `=SMILES2EXACTMASS(A2)` | `=INCHI2EXACTMASS(A2)` · `=KET2EXACTMASS(A2)` | `180.0423` | **0ms 오프라인** · 단일동위원소 질량 `Double` (Da, 소수 4자리) |
| **물질 일반명 (Name)** | `=SMILES2NAME(A2)` | `=INCHI2NAME(A2)` · `=KET2NAME(A2)` | `Aspirin` | PubChem / IUPAC 조회 + 로컬 영구 캐시 (`String`) |

## 2. 실전 템플릿 ①: 유기 합성 당량(ELN) 시약 칭량 계산기

`=SMILES2MW` 는 실제 숫자(`Double`)를 반환하므로 `mmol` 셀과 바로 곱해 시약 칭량값(`mg`)과 이론 수율을 자동 계산할 수 있습니다:

| A열 (시약명) | B열 (`=NAME2SMILES(A2)`) | C열 (`=SMILES2IMG(B2)`) | D열 (`=SMILES2MW(B2)`) | E열 (당량) | F열 (`mmol`) | G열 (`=D2*F2` mg) |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D 구조식)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D 구조식)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |

생성물 행의 SMILES(`B4`)에 대해서도 한계 시약 `mmol`(`$F$2`) × `SMILES2MW(B4)` 수식을 넣으면 **이론 수율**(`mg`)이 즉시 계산됩니다.

## 3. 실전 템플릿 ②: LC-MS / HRMS 이온 피크 계산표 (*m/z*)

소수점 4자리까지 계산되는 `=SMILES2EXACTMASS(A2)` 를 이용해 질량분석(HRMS) 이론 이온 피크를 자동 계산해 두세요:

| 질량분석 이온 (Adduct) | 복사해서 쓰는 수식 (`A2` = 구조 셀) | 아스피린 (`C9H8O4`) |
| :--- | :--- | :---: |
| **중성 정밀질량 (`M`)** | `=SMILES2EXACTMASS(A2)` | `180.0423` |
| **`[M + H]+`** *(ESI 양이온)* | `=ROUND(SMILES2EXACTMASS(A2) + 1.007276, 4)` | `181.0496` |
| **`[M + NH4]+`** | `=ROUND(SMILES2EXACTMASS(A2) + 18.033823, 4)` | `198.0761` |
| **`[M + Na]+`** | `=ROUND(SMILES2EXACTMASS(A2) + 22.989218, 4)` | `203.0315` |
| **`[M - H]-`** *(ESI 음이온)* | `=ROUND(SMILES2EXACTMASS(A2) - 1.007276, 4)` | `179.0350` |

## 관련 가이드

- [**Lipinski · ADME · 3D**](/ko/excel/lipinski-adme) — 경구 흡수성, CNS BBB 기준, 3D 복잡도(`Fsp3`) 및 `IF`/`AND` 선별식
- [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) — 시약명·CAS·표준 SMILES·InChIKey 상호 변환
- [**예제 파일 (.xlsx)**](/ko/excel/playbook) — `2. Properties & Lipinski` · `7. Real-World Workflows` 탭에서 라이브 템플릿 복사
