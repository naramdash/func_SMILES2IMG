# Excel 레시피 1부: 구조식 표 · 이름/CAS · 스타일 · 반응식

> **Excel 가이드 구성 방식 — 실무 레시피 북:** 모든 레시피는 완성 화면을 먼저 보여주고, **Microsoft 365 / 2024**와 **구버전 Excel 2016–2021(`.FLOAT`)** 수식을 함께 안내하며, 실제 구동 스크린샷과 사용 영상을 함께 제공합니다.

<FeatureMedia id="excel-overview" />

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 8개 탭 마스터 쇼케이스 엑셀 파일 다운로드 (.xlsx)</div>
    <p class="s2i-card-desc">0.빠른시작 · 1.구조식생성 · 2.물성/Lipinski/3D · 3.부분구조/유사도 · 4.스타일매트릭스 · 5.반응식 · 6.구형호환 · 7.리본도구</p>
  </a>
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 다음 편: 레시피 6~10 (물성 · Lipinski · SAR · SDF) ➔</div>
    <p class="s2i-card-desc">분자량, 약물성(Ro5), 부분구조 검색, 유사도 순위 및 .sdf 입출력 레시피</p>
  </a>
</div>

## Excel 입력 방식 3가지 & 버전 선택

| 입력 방식 | 동작 방식 | 추천 상황 |
| :--- | :--- | :--- |
| **1. 셀 수식** | `=SMILES2IMG(A2)`, `=NAME2IMG("Aspirin")`, `=KET2IMG(A2)` | 수백 행의 화합물 표를 정렬·필터와 연동해 일괄 생성할 때 |
| **2. 상단 `[SMILES]` 리본** | **`[Insert / Edit (SMILES)]`** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd>) 또는 **`[Paste Structure]`** | 마우스로 직접 그리거나 복사해 둔 SMILES를 즉시 넣을 때 |
| **3. 마우스 조작** | 셀 **더블클릭** ➔ 제자리 수정 / **우클릭** ➔ `Chemical Structure` | 기존 구조 수정, SVG/물성 복사, 상세 정보 창 열기 |

::: info 💡 내 엑셀 버전에 맞는 함수 선택 (365 vs `.FLOAT`)
* **Microsoft 365 / Excel 2024 이상:** `=SMILES2IMG(B2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(C2)` — 정렬·필터 시 행과 함께 움직이는 **네이티브 인셀 이미지**로 삽입됩니다.
* **Excel 2016 / 2019 / 2021:** 함수 이름 뒤에 **`.FLOAT`**만 붙이세요(`=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)`) — 셀 크기와 이동에 자동 연동되는 그림 도형(`xlMoveAndSize`)으로 생성됩니다.
:::

---

## 레시피 1: SMILES 목록으로 고화질 구조식 표 만들기

**목표:** A열에 화합물명, B열에 SMILES를 모아두고 C열에 `=SMILES2IMG(B2)` 수식을 채워 비율이 유지되는 구조식 표를 만듭니다.

```text
=SMILES2IMG(B2)
```

<FeatureMedia id="excel-formula-table" />

1. `A2:A5`에 화합물 이름을, `B2:B5` 셀에 SMILES 문자열(예: 아스피린 `CC(=O)Oc1ccccc1C(=O)O`)을 입력합니다.
2. `C2` 셀에 `=SMILES2IMG(B2)`를 입력하고 <kbd>Enter</kbd>를 누릅니다.
3. **행 높이(`62~80 pt`)**와 **열 너비**를 넓히면 비율이 깨지지 않고 자동으로 커집니다.
4. 채우기 핸들을 `C2:C5` 아래로 드래그해 수백 행을 한 번에 채웁니다.

---

## 레시피 2: 시약 이름 · 실험실 약어 · CAS 번호로 구조식 바로 그리기

**목표:** SMILES를 따로 찾지 않고 영문 물질명, 실험실 약어(`THF`, `DMSO`, `DCM`), 또는 CAS 등록번호만으로 구조식을 그립니다.

```text
=NAME2IMG("Aspirin")
=NAME2IMG("50-78-2")
=NAME2SMILES("Ibuprofen")
```

<FeatureMedia id="excel-name-cas" />

::: details 📚 인터넷 없이 0ms 즉시 변환되는 내장 약어 · 시약 · CAS 사전
주요 실험실 용매, 의약품, 생체분자와 CAS 번호가 내장되어 **오프라인에서도 대기 시간 없이(0ms) 즉시 변환**됩니다. 목록에 없는 화합물은 최초 1회 PubChem에서 조회한 뒤 내 PC에 영구 캐싱됩니다:
* **실험실 용매·약어:** `THF` (`109-99-9`), `DMSO` (`67-68-5`), `DMF` (`68-12-2`), `DCM` (`75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone`, `Ethanol`, `Methanol`, `Pyridine`, `Toluene`, `Benzene`
* **의약품·천연물:** `Aspirin` (`50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen`, `Penicillin G`, `Amoxicillin`, `Artemisinin`, `Paclitaxel`
* **생체분자:** `Glucose`, `Vitamin C`, `Citric Acid`, `Glycine`, `Alanine`, `Dopamine`, `Serotonin`, `Epinephrine`
:::

::: details ⚖️ 합성 실험노트(ELN)용 시약 당량 · 칭량 무게(mg) 자동 계산 표 (`=SMILES2MW`)
`=SMILES2MW`는 텍스트가 아닌 **실제 숫자(`double`)**를 반환하므로 몰수(`mmol`)와 곱해 무게(`mg`)를 바로 구할 수 있습니다:
| A열 (시약명) | B열 (`=NAME2SMILES(A2)`) | C열 (`=SMILES2IMG(B2)`) | D열 (`=SMILES2MW(B2)`) | E열 (당량) | F열 (`mmol`) | G열 (`=D2*F2` mg) |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| `Salicylic Acid` | `O=C(O)c1ccccc1O` | *(2D 구조식)* | `138.12` | `1.0` | `5.0` | **`690.6 mg`** |
| `Acetic Anhydride` | `CC(=O)OC(=O)C` | *(2D 구조식)* | `102.09` | `1.5` | `=$F$2*E3` (`7.5`) | **`765.7 mg`** |
:::

---

## 레시피 3: 논문용 흑백(ACS) · 작용기 강조 · 다크 모드 · 원자 번호 스타일

**목표:** 두 번째(`[배경색]`)와 세 번째(`[스타일]`) 인수에 키워드를 조합해 용도별 전문 도면을 만듭니다.

| 원하는 스타일 | 수식 예시 | 활용처 |
| :--- | :--- | :--- |
| **기본 컬러 (투명 배경)** | `=SMILES2IMG(A2)` | 일반 보고서 (O 빨강, N 파랑) |
| **학술지 투고용 흑백 (ACS)** | `=SMILES2IMG(A2, "", "bw")` | 논문·특허용 순수 흑백 선화 |
| **특정 골격 하이라이트** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | 일치하는 SMARTS 작용기 컬러 강조 |
| **다크 모드 (흰색 결합선)** | `=SMILES2IMG(A2, "#1E293B", "white")` | 어두운 셀 배경 위 선명한 흰색 선 |
| **NMR 원자 번호 & 입체** | `=SMILES2IMG(A2, "", "num stereo")` | 1부터 시작하는 원자 번호 및 `R/S`, `E/Z` 표기 |

<FeatureMedia id="excel-style-matrix" />

> 🛠️ 옵션을 마우스로 조합해 보려면 [**수식 생성기 & 스타일 시뮬레이터**](/ko/tools/formula-builder)를 열어보세요.

---

## 레시피 4: 반응 화살표·촉매가 포함된 화학 반응식 (`=KET2IMG` + Ketcher)

**목표:** 반응물, 생성물, 화살표(`➔`), `+`, 시약 주석을 무손실 KET JSON으로 셀에 저장하고 언제든 다시 편집합니다.

```text
=KET2IMG(A2)
```

<FeatureMedia id="excel-reactions-ket" />

1. 셀을 선택하고 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르거나 리본 **`[Insert / Edit (KET)]`**를 클릭합니다.
2. 오프라인 **Ketcher 2D 편집기**에서 반응 화살표와 촉매·조건 텍스트를 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 `=KET2IMG(...)` 수식으로 삽입됩니다. 이후 셀을 **더블클릭**하면 0.05초 만에 원본 반응식이 그대로 다시 열립니다.

---

## 레시피 5: 구버전 Excel 2016–2021 호환 모드 (`.FLOAT`)

**목표:** 인셀 이미지 함수를 지원하지 않는 구버전 Excel에서도 셀 크기와 이동에 맞춰 따라 움직이는 구조식을 사용합니다.

```text
=SMILES2IMG.FLOAT(A2)
=NAME2IMG.FLOAT("Caffeine")
```

<FeatureMedia id="excel-legacy-float" />

생성된 부동 그림 도형은 셀 위치와 크기 변경(`xlMoveAndSize`)에 자동으로 연동됩니다.

---

## Excel `[SMILES]` 상단 리본 & 우클릭 메뉴 요약

<FeatureMedia id="excel-ribbon-tools" />

* **`Insert / Edit (SMILES / InChI / KET)` (`Ctrl+Shift+K`):** 선택한 셀에 2D 편집기를 열어 원하는 포맷의 수식을 생성합니다.
* **`3D Structure (Mol*)`:** 3D 단백질 뷰어를 실행합니다 (`1CRN` 같은 4자리 PDB 코드가 셀에 있으면 자동 로드).
* **`Paste Structure`:** 클립보드나 셀 텍스트를 0.1초 만에 구조식 수식으로 바꾸고 행·열 크기를 자동 조절합니다.
* **`Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** 식별자 텍스트 또는 Illustrator용 무손실 벡터 SVG를 복사합니다.
* **`Molecule Info`:** 17가지 물성 상세 창을 열거나 TSV·마크다운·JSON 표 서식으로 복사합니다.
* **`Batch Style`:** 선택한 범위(또는 활성 시트 전체)를 클릭 한 번으로 컬러·흑백(ACS)·흰 배경·다크 모드로 통일합니다.
* **`SDF File`:** `.sdf` 파일을 12열 약물성 시트로 불러오거나 시트를 `.sdf`로 내보냅니다.
* **`Diagnostics` & `About func_SMILES2IMG`:** `#VALUE!` 파서 원인 진단 로그를 보거나 업데이트를 확인합니다.

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Excel 실제 3단 Chemical Structure 우클릭 메뉴" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🚀 Excel 가이드 이어서 보기

<div class="s2i-card-grid">
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel 레시피 6~10: 물성 · SAR · SDF</div>
    <p class="s2i-card-desc">분자량, 정밀질량(HRMS), Lipinski Ro5, 12대 SMARTS 작용기 검색 및 .sdf 입출력</p>
  </a>
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📑 40개 전체 엑셀 수식 사전</div>
    <p class="s2i-card-desc">모든 함수 시그니처, 반환 타입, 스타일 키워드 별칭 및 오류 코드 안내</p>
  </a>
</div>
