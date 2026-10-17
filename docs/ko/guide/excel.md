# Excel 기초: 구조식 표 만들기 & 스타일링

엑셀 시트 안에서 수백 개의 화학 구조식을 실시간으로 그리고, 원하는 스타일(흑백, 하이라이트, 다크 모드)로 꾸미며, 셀 더블클릭으로 언제든 수정하는 방법을 안내합니다.

> 🔬 **분자량·정밀질량 계산, 약물성(Lipinski), 부분구조 검색(SAR), `.sdf` 파일 연동**을 찾으시나요?  
> 👉 [**Excel 심화: 물성 계산 · 신약 스크리닝 · SDF 연동 가이드 ➔**](/ko/guide/excel-cheminformatics)에서 바로 확인하세요!

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 엑셀 핵심 사용법 4줄 요약
1. **수식으로 구조식 넣기:** 셀에 `=SMILES2IMG(A2)` 또는 `=NAME2IMG("Aspirin")`을 입력하고 아래로 드래그 채우기.
2. **직접 그리기:** 빈 셀에서 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 눌러 분자나 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
3. **더블클릭 제자리 수정:** 구조식이 들어있는 셀을 **더블클릭**하면 언제든 편집기가 다시 열려 바로 고칠 수 있음.
4. **우클릭 물성 복사:** 셀을 **우클릭** ➔ **`Chemical Structure`** 메뉴에서 분자량, 화학식, SVG 벡터를 원클릭 복사.
:::

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

## 2. 클릭으로 조합해 보는 스타일 옵션 (흑백 · 하이라이트 · 다크모드)

아래 위젯에서 원하는 옵션을 클릭해 보고, 생성된 수식을 그대로 복사해 엑셀에 붙여넣으세요!

<FormulaBuilder />

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
