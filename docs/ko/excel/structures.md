# 구조식 표 · 이름/CAS

SMILES 문자열, 화합물 이름, 실험실 용매 약어, CAS 번호가 담긴 열을 셀 크기에 맞춰 선명하게 확대되는 2D 구조식 표로 자동 변환합니다.

## 1. SMILES 열을 구조식 표로 변환 (`=SMILES2IMG`)

A열에 화합물 이름, B열에 SMILES를 두고 C열에 수식 한 줄로 구조식을 자동 배치합니다:

```text
=SMILES2IMG(B2)
```

<FeatureMedia id="excel-formula-table" />

1. `A2:A5`에 화합물 이름, `B2:B5`에 SMILES 문자열(예: 아스피린 `CC(=O)Oc1ccccc1C(=O)O`)을 입력합니다.
2. `C2` 셀에 `=SMILES2IMG(B2)`를 입력하고 <kbd>Enter</kbd>를 누릅니다.
3. **행 높이**(`62~80 pt`)와 **열 너비**를 넉넉히 조절하면 비율이 유지된 채 선명하게 확대됩니다.
4. 채우기 핸들을 아래로 드래그해 수백 행을 한 번에 렌더링합니다.

## 2. 화합물 이름 · 용매 약어 · CAS 번호 직접 렌더링 (`=NAME2IMG`)

SMILES를 따로 찾을 필요 없이 영문 물질명, 용매 약어(`THF`, `DMSO`, `DCM`), CAS 번호를 바로 넘깁니다:

```text
=NAME2IMG("Aspirin")
=NAME2IMG("50-78-2")
=NAME2SMILES("Ibuprofen")
```

<FeatureMedia id="excel-name-cas" />

::: details 내장 0ms 오프라인 사전 (용매 · 의약품 · 생체분자 · CAS)
자주 쓰는 용매·의약품·생체분자·CAS 번호는 **인터넷 없이 0ms 만에 로컬 변환**됩니다. 사전에 없는 화합물은 PubChem에서 1회 조회 후 디스크에 로컬 캐시됩니다:
* **실험실 용매:** `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`
* **의약품:** `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel`
* **생체분자:** `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`
:::

::: details `=SMILES2MW`를 활용한 당량(ELN) 계산 표 예시
`=SMILES2MW`는 실제 숫자(`double`)를 반환하므로 `mmol` 셀과 바로 곱해 시약 칭량값(`mg`)을 계산할 수 있습니다:

| A열 (시약명) | B열 (`=NAME2SMILES(A2)`) | C열 (`=SMILES2IMG(B2)`) | D열 (`=SMILES2MW(B2)`) | E열 (당량) | F열 (`mmol`) | G열 (`=D2*F2` mg) |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D 구조식)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D 구조식)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |
:::
