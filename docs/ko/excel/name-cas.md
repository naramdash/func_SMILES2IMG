# 시약명 · 용매 · CAS 구조식 렌더링

SMILES 문자열을 따로 찾지 않아도 **영문 화합물 이름**, **실험실 용매·커플링 시약 약어**(`THF`, `DMF`, `HATU`), **CAS 등록번호**(`"50-78-2"`)만으로 `=NAME2IMG` 를 통해 셀 안에 2D 화학 구조식을 즉시 렌더링합니다.

<FormulaTable
  title="시약명 · 용매 · CAS 시트"
  :columns='["시약명 / CAS", "수식"]'
  :rows='[
    { input: "Aspirin", formula: "=NAME2IMG(A2)", desc: "영문 시약명·의약품 이름으로 즉시 렌더링", out: "아스피린 2D" },
    { input: "50-78-2", formula: "=NAME2IMG(A3)", desc: "CAS 등록번호 직접 입력 렌더링", out: "아스피린 2D" },
    { input: "", formula: "=NAME2IMG(\"THF\", \"\", \"bw\")", desc: "실험실 용매 약어 ACS 흑백 선화 단독 출력", out: "흑백 THF" }
  ]'
/>

<FeatureMedia id="excel-name-cas" />

## 1. 0ms 내장 오프라인 사전 + 로컬 PubChem 영구 캐시

`=NAME2IMG(이름_또는_CAS, [배경색], [스타일], [변형])` (및 Excel 2016~2021용 `=NAME2IMG.FLOAT`)는 먼저 내장 오프라인 사전을 0ms 속도로 조회합니다. 내장 사전에 없는 신규 화합물만 PubChem에서 최초 1회 조회한 뒤 디스크 캐시에 저장하므로, 이후에는 오프라인 상태에서도 즉시 재렌더링됩니다.

| 분류 | 0ms 내장 오프라인 지원 항목 (주요 예시) | 수식 예시 |
| :--- | :--- | :--- |
| **실험실 용매** | `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`, `Dioxane`, `MTBE` | `=NAME2IMG("THF")` |
| **커플링 시약 · 염기** | `TEA`, `DIPEA`, `DMAP`, `EDCI`/`EDC`, `HOBt`, `HATU`, `TFA`, `Boc2O`, `Fmoc-OSu`, `NaBH4`, `LiAlH4`, `mCPBA` | `=NAME2IMG("HATU")` |
| **주요 의약품** | `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel` | `=NAME2IMG("50-78-2")` |
| **생체분자 · 아미노산** | `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`, `Cholesterol` | `=NAME2IMG("Caffeine")` |

::: tip 공용 `[배경색]` · `[스타일]` 옵션을 100% 그대로 지원합니다
`=NAME2IMG` 역시 `=SMILES2IMG`와 완전히 동일한 옵션 파서를 사용합니다:
* 논문 투고용 ACS 흑백: `=NAME2IMG(A2, "", "bw")`
* 다크 모드 흰색 선화: `=NAME2IMG(A2, "#1E293B", "white")`
* 카복실산 작용기 강조: `=NAME2IMG(A2, "", "hl:C(=O)[OH]")`

전체 옵션 키워드는 [**스타일 · 하이라이트**](/ko/excel/styles-reactions) 페이지에서 확인하세요.
:::

## 2. `=NAME2SMILES` 및 분자량(`MW`) 열과 함께 쓰기

실험실 시약 재고표나 합성 노트(ELN)에서는 A열에 시약명이나 CAS 번호만 적고, 옆 열에서 SMILES 텍스트·구조식 이미지·분자량을 동시에 자동 채움할 수 있습니다:

| A열 (시약명 / CAS 입력) | B열 (`=NAME2SMILES(A2)`) | C열 (`=NAME2IMG(A2)`) | D열 (`=SMILES2MW(B2)`) |
| :--- | :--- | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D 구조식)* | `138.12` |
| `50-78-2` *(Aspirin)* | `CC(=O)Oc1ccccc1C(=O)O` | *(2D 구조식)* | `180.16` |
| `DIPEA` | `CCN(C(C)C)C(C)C` | *(2D 구조식)* | `129.24` |

## 관련 가이드

- [**SMILES · InChI**](/ko/excel/structures) — 공용 4인수 구조 및 대량 SMILES/InChI 구조식 표 만들기
- [**스타일 · 하이라이트**](/ko/excel/styles-reactions) — 배경색, ACS 흑백, SMARTS 강조, 회전·반전 키워드 전체 규격
- [**표기법 변환 · InChIKey**](/ko/excel/data-sdf) — 시약명·CAS를 표준 SMILES(`=NAME2SMILES`) 및 InChIKey로 변환하기
