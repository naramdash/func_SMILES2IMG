# Excel 1부: 구조식 테이블 & 이름/CAS 변환

시약 이름, CAS 번호, 또는 SMILES 문자열만으로 엑셀 셀 안에 수백 개의 고화질 구조식 표를 자동 생성하고, 셀을 더블클릭해 그 자리에서 바로 수정하는 실무 가이드입니다.

<p align="center">
  <img src="/excel-showcase.png" alt="func_SMILES2IMG Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 수식이 입력된 샘플 엑셀 파일 다운로드 (.xlsx)</div>
    <p class="s2i-card-desc">직접 타이핑할 필요 없이, 구조식과 스타일 수식이 미리 적용된 시트를 바로 열어보세요</p>
  </a>
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 이어서 보기: Excel 2부 (물성·Lipinski·SAR·SDF) ➔</div>
    <p class="s2i-card-desc">옆 열에 분자량(MW), 약물성(Ro5), 부분구조 검색을 붙이고 .sdf 파일을 입출력하세요</p>
  </a>
</div>

## ⚡ 복사해서 바로 쓰는 대표 수식 3선

빈 엑셀 셀에 아래 수식을 그대로 복사해 붙여넣어 보세요 (구버전 Excel 2016~2021은 함수명 뒤에 **`.FLOAT`**만 붙이면 됩니다):

```text
=NAME2IMG("Aspirin")
```
```text
=SMILES2IMG(A2)
```
```text
=SMILES2IMG(A2, "", "bw")
```

---

## 1. 수식으로 화합물 표 & 물성 자동 채우기

엑셀에서 가장 많이 쓰는 방법은 A열에 SMILES(또는 시약 이름/CAS 번호)를 적어두고, 옆 열에 수식을 걸어 아래로 드래그하는 것입니다.

### 실전 시트 구성 예시

| 열 (Column) | 항목 | B2 ~ E2 셀에 입력할 수식 | 결과 예시 (`A2` = 아스피린) |
| :---: | :--- | :--- | :--- |
| **A열** | 입력값 (SMILES 또는 이름) | `CC(=O)Oc1ccccc1C(=O)O` | 원본 텍스트 |
| **B열** | **2D 화학 구조식** | `=SMILES2IMG(A2)` | *(셀 크기에 맞춘 구조식 그림)* |
| **C열** | **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **D열** | **분자량 (MW, g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **E열** | **약물성 (Lipinski Ro5)** | `=SMILES2LIPINSKI(A2)` | `PASS` |

* **시약 이름, 실험실 약어(`THF`, `DMSO`, `DCM`), 또는 CAS 번호만 알고 있다면?**  
  `=NAME2IMG("Aspirin")`, `=NAME2IMG("THF")`, `=NAME2IMG("50-78-2")`처럼 바로 그림을 그리거나, `=NAME2SMILES("Ibuprofen")`으로 SMILES 문자열을 먼저 구할 수 있습니다.
* **그림을 더 크게 보고 싶다면?**  
  엑셀의 **행 높이**(예: 80~100 pt)와 **열 너비**를 넓히면 비율이 깨지지 않고 자동으로 선명하게 커집니다.

::: details ⚖️ 합성 실험노트(ELN)용 시약 당량 · 칭량 무게(mg) 자동 계산 표 만들기 (클릭하여 펼치기)
`=SMILES2MW(B2)` 함수는 텍스트가 아닌 **실제 숫자**(Number)를 반환하므로, 엑셀 사칙연산(`*`, `/`)과 바로 결합하여 **유기 합성 반응의 시약 당량(Eq) 및 저울로 달아야 할 무게(mg) 자동 계산 표**를 만들 수 있습니다 (`E2` 셀에 기준 물질 몰수 `5.0 mmol` 입력 기준):

| 열 | 항목 | 2행 (출발 물질: Salicylic Acid) | 3행 (시약: Acetic Anhydride) | 설명 |
| :---: | :--- | :--- | :--- | :--- |
| **A열** | **시약 이름** | `Salicylic Acid` | `Acetic Anhydride` | 시약명 또는 CAS 번호 |
| **B열** | **SMILES** | `=NAME2SMILES(A2)` | `=NAME2SMILES(A3)` | 이름에서 SMILES 자동 추출 |
| **C열** | **구조식** | `=SMILES2IMG(B2)` | `=SMILES2IMG(B3)` | 2D 구조식 자동 렌더링 |
| **D열** | **분자량 (g/mol)** | `=SMILES2MW(B2)` ➔ `138.12` | `=SMILES2MW(B3)` ➔ `102.09` | **숫자 타입**으로 즉시 반환 |
| **E열** | **당량 (Eq)** | `1.0` | `1.5` | 실험자가 입력하는 당량비 |
| **F열** | **몰수 (mmol)** | `5.0` *(기준 몰수)* | `=$F$2 * E3` ➔ `7.5` | 기준 몰수 × 당량 자동 계산 |
| **G열** | **칭량할 무게 (mg)** | `=D2 * F2` ➔ **`690.6 mg`** | `=D3 * F3` ➔ **`765.7 mg`** | **분자량(D열) × 몰수(F열)** 자동 계산! |
:::

::: details 📚 인터넷 없이 0.001초 만에 즉시 변환되는 내장 약어·용매·약물 사전 보기
아래의 다빈도 실험실 용매 약어, 주요 의약품, 생체분자 및 해당 CAS 번호는 프로그램 내부에 기본 내장되어 있어 **오프라인 상태에서도 대기 시간 없이(0ms) 즉시 변환**됩니다. 목록에 없는 수만 종의 기타 시약 이름이나 CAS 번호는 최초 1회 PubChem에서 자동 조회한 뒤 내 PC에 영구 저장(캐시)되어 다음부터는 오프라인에서도 즉시 작동합니다.

| 분류 | 바로 입력 가능한 영문 이름 · 약어 · CAS 번호 예시 |
| :--- | :--- |
| **실험실 필수 용매 & 약어** | `THF` (`Tetrahydrofuran`, `109-99-9`), `DMSO` (`Dimethyl Sulfoxide`, `67-68-5`), `DMF` (`Dimethylformamide`, `68-12-2`), `DCM` (`Dichloromethane`, `75-09-2`), `Acetonitrile` (`75-05-8`), `Ethyl Acetate` (`141-78-6`), `Acetone` (`67-64-1`), `Ethanol` (`64-17-5`), `Methanol` (`67-56-1`), `Isopropanol` (`67-63-0`), `Diethyl Ether` (`60-29-7`), `Chloroform` (`67-66-3`), `Hexane` (`110-54-3`), `Cyclohexane`, `Benzene`, `Toluene`, `Phenol`, `Aniline`, `Pyridine`, `Acetic Acid`, `Formic Acid`, `Benzoic Acid`, `Water` |
| **대표 의약품 & 천연물** | `Aspirin` (`Acetylsalicylic acid`, `50-78-2`), `Caffeine` (`58-08-2`), `Acetaminophen` (`Paracetamol`, `Tylenol`, `103-90-2`), `Ibuprofen` (`15687-27-1`), `Naproxen` (`22204-53-1`), `Penicillin G` (`61-33-6`), `Amoxicillin` (`26787-78-0`), `Artemisinin` (`63968-64-9`), `Paclitaxel` (`Taxol`, `33069-62-4`), `Salicylic Acid` (`69-72-7`), `Cholesterol` (`57-88-5`) |
| **생체분자 · 비타민 · 아미노산** | `Glucose` (`D-Glucose`, `50-99-7`), `Vitamin C` (`Ascorbic Acid`, `50-81-7`), `Citric Acid` (`77-92-9`), `Lactic Acid` (`50-21-5`), `Glycine` (`56-40-6`), `Alanine` (`56-41-7`), `Dopamine` (`51-61-6`), `Serotonin` (`50-67-9`), `Epinephrine` (`Adrenaline`, `51-43-4`) |
:::

::: info 💡 내 엑셀 버전에 맞는 이미지 함수 고르기
* **Microsoft 365 / Excel 2024 이상:** `=SMILES2IMG(A2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`  
  *(셀 내부에 진짜 인셀 이미지로 들어가며 정렬·필터 시 데이터와 함께 움직입니다. 참고로 `=SMILES2IMG(A2)` 하나만 써도 `A2` 셀의 SMILES, InChI, KET 데이터를 모두 자동 인식합니다!)*
* **Excel 2016 / 2019 / 2021:** 함수 이름 뒤에 <code><strong>.FLOAT</strong></code>만 붙이세요!  
  예: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)` *(셀 테두리에 딱 맞춰 함께 이동하고 크기가 조절되는 그림 도형으로 삽입됩니다)*
:::

---

## 2. 자주 쓰는 스타일 레시피 (흑백 · 하이라이트 · 다크모드)

두 번째 인수(`[배경색]`)와 세 번째 인수(`[스타일]`)에 간단한 키워드만 넣으면 원하는 스타일로 바뀝니다:

| 원하는 스타일 | 복사해서 바로 쓸 수식 (`A2` 기준) | 설명 |
| :--- | :--- | :--- |
| **🎨 기본 컬러 (투명 배경)** | `=SMILES2IMG(A2)` | 산소(빨강), 질소(파랑) 표준 원소 색상 |
| **📄 학술지 투고용 흑백 (ACS)** | `=SMILES2IMG(A2, "", "bw")` | 모든 원자와 결합을 깔끔한 검은색 선화로 통일 |
| **🔍 특정 골격 하이라이트** | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` | 지정한 부분구조(예: 벤젠 고리)를 붉은색으로 강조 |
| **🌙 다크 모드 (어두운 셀)** | `=SMILES2IMG(A2, "#1a1a1a", "white")` | 짙은 배경 위에 선명한 흰색 결합선으로 렌더링 |
| **🔢 NMR 분석용 원자 번호** | `=SMILES2IMG(A2, "", "num stereo")` | 모든 원자에 1번부터 인덱스 번호 및 입체(`R/S`) 표시 |

> 🛠️ **클릭만으로 옵션을 조합해 보고 싶으신가요?**  
> 👉 [**엑셀 함수 표현 생성기 & 스타일 시뮬레이터 열기 ➔**](/ko/tools/formula-builder)에서 마우스 클릭으로 실시간 미리보기를 확인하고 완성된 수식을 복사하세요!

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 3. 마우스로 직접 그리기 & 화학 반응식 (`Ctrl + Shift + K`)

SMILES 코드를 모르거나 다단계 유기 합성 반응식을 넣고 싶을 때는 내장된 **Ketcher 2D 화학 편집기**를 사용하세요.

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. 구조식을 넣을 셀을 클릭하고 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누릅니다 (또는 상단 **`[SMILES]` ➔ `[Insert / Edit (SMILES)]`** 클릭).
2. 벤젠 고리, 작용기, 입체 결합(쐐기/점선), 또는 **반응 화살표와 촉매 텍스트**(H₂SO₄, 55 °C)를 자유롭게 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 셀에 즉시 삽입됩니다!
   * 단일 분자는 `=SMILES2IMG(...)` 수식으로 들어가고,
   * 반응 화살표나 텍스트가 포함된 반응식은 위치와 글자가 100% 보존되는 `=KET2IMG(...)` 수식으로 자동 저장됩니다.
4. **다시 고치려면?** 해당 셀을 **마우스로 더블클릭**하면 그렸던 분자나 반응식이 0.05초 만에 그대로 다시 열립니다.

---

## 4. 상단 `[SMILES]` 리본 메뉴 & 우클릭 메뉴 요약

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Office" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ `Insert / Edit (SMILES)` (`Ctrl+Shift+K`):** 2D 분자 스케처 실행 (옆의 `InChI`, `KET` 버튼으로 저장 수식 형태를 직접 지정할 수도 있습니다)
* **🧬 `3D Structure (Mol*)`:** 3D 단백질 뷰어 실행 (현재 선택한 셀에 `1CRN` 같은 4자리 PDB 코드가 적혀 있으면 자동으로 해당 단백질을 불러옵니다!)
* **📋 `Paste Structure`:** 클립보드에 복사해 둔 SMILES나 현재 셀에 적힌 SMILES 텍스트를 0.1초 만에 구조식 수식으로 변환 (셀 행 높이·열 너비도 보기 좋게 자동 확대됩니다!)
* **🖼️ `Copy SMILES` / `Copy InChI` / `Copy KET` / `Copy SVG`:** 선택한 분자에서 화학 식별자 텍스트나 Illustrator용 무손실 벡터 그래픽(`SVG`)을 원클릭 복사
* **🔬 `Molecule Info`:** 물질명·화학식·분자량·LogP·Lipinski 약물성 확인 및 엑셀 행(TSV)·마크다운 표 복사
* **🎨 `Batch Style`:** 시트 전체 구조식을 컬러 / 논문용 흑백(ACS) / 흰색 배경 / 다크 모드로 일괄 변환
* **📂 `SDF File`:** `.sdf` / `.mol` 파일을 8열 엑셀 표로 불러오거나(`Import`) 시트 데이터를 `.sdf`로 내보내기(`Export`)

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: details 🔍 수식이 #VALUE! 오류를 반환할 때 (Diagnostics 진단 창 활용법)
참조한 셀의 SMILES 문법이 잘못되었을 때(예: 괄호나 고리 번호가 안 닫힌 경우) `#VALUE!`가 표시됩니다.  
상단 리본의 **`[SMILES]` ➔ `[Diagnostics]`** 버튼을 클릭하면 어떤 셀의 SMILES에서 무슨 문법 오류가 났는지 정확한 화학 파서 원인을 실시간 로그로 확인할 수 있습니다.
:::

---

## 🚀 다음 단계로 이어서 보기

<div class="s2i-card-grid">
  <a href="../guide/excel-cheminformatics" class="s2i-card">
    <div class="s2i-card-title">🔬 Excel 심화: 물성 · SAR · SDF</div>
    <p class="s2i-card-desc">분자량, 정밀질량(HRMS), Lipinski 약물성, 12대 작용기 검색 사전 및 .sdf 파일 연동</p>
  </a>
  <a href="../guide/publishing" class="s2i-card">
    <div class="s2i-card-title">🎨 출판 스타일 · SVG · 물성 복사</div>
    <p class="s2i-card-desc">시트 전체 ACS 흑백 일괄 변환, Illustrator 벡터 복사, 5가지 표 서식 내보내기</p>
  </a>
</div>
