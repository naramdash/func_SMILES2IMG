# Microsoft Excel 가이드

엑셀 시트 안에서 수백 개의 화학 구조식을 실시간으로 그리고, 분자량·화학식·약물성(Lipinski)을 자동 계산하며, `.sdf` 화학 데이터 파일을 자유롭게 주고받는 방법을 안내합니다.

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 엑셀 핵심 사용법 4줄 요약
1. **수식으로 구조식 넣기:** 셀에 `=SMILES2IMG(A2)` 또는 `=NAME2IMG("Aspirin")`을 입력하고 아래로 드래그 채우기.
2. **직접 그리기:** 빈 셀에서 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 눌러 분자나 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>.
3. **더블클릭 제자리 수정:** 구조식이 들어있는 셀을 **더블클릭**하면 언제든 편집기가 다시 열려 바로 고칠 수 있음.
4. **우클릭 물성 복사:** 셀을 **우클릭 ➔ `Chemical Structure`**에서 분자량, 화학식, SVG 벡터를 원클릭 복사.
:::

---

## 1. 수식으로 화합물 표 & 물성 자동 채우기

엑셀에서 가장 많이 쓰는 방법은 A열에 SMILES(또는 시약 이름/CAS 번호)를 적어두고, 옆 열에 수식을 걸어 아래로 드래그하는 것입니다.

### 실전 시트 구성 예시

| 열 (Column) | 항목 |B2 ~ E2 셀에 입력할 수식 | 결과 예시 (A2 = `"Aspirin"` 또는 SMILES) |
| :---: | :--- | :--- | :--- |
| **A열** | 입력값 (SMILES 또는 이름) | `CC(=O)Oc1ccccc1C(=O)O` | 원본 텍스트 |
| **B열** | **2D 화학 구조식** | `=SMILES2IMG(A2)` | *(셀 크기에 맞춘 구조식 그림)* |
| **C열** | **화학식 (Formula)** | `=SMILES2FORMULA(A2)` | `C9H8O4` |
| **D열** | **분자량 (MW, g/mol)** | `=SMILES2MW(A2)` | `180.16` |
| **E열** | **약물성 (Lipinski Ro5)** | `=SMILES2LIPINSKI(A2)` | `PASS` |

* **시약 이름이나 CAS 번호만 알고 있다면?**  
  `=NAME2IMG("Aspirin")`, `=NAME2IMG("50-78-2")`처럼 바로 그림을 그리거나, `=NAME2SMILES("Ibuprofen")`으로 SMILES 문자열을 먼저 구할 수 있습니다.
* **그림을 더 크게 보고 싶다면?**  
  엑셀의 **행 높이(예: 80~100 pt)**와 **열 너비**를 넓히면 비율이 깨지지 않고 자동으로 선명하게 커집니다.

::: info 💡 내 엑셀 버전에 맞는 이미지 함수 고르기
* **Microsoft 365 / Excel 2024 이상:** `=SMILES2IMG(A2)`, `=NAME2IMG(A2)`, `=KET2IMG(A2)`, `=INCHI2IMG(A2)`  
  *(셀 내부에 진짜 인셀 이미지로 들어가며 정렬·필터 시 데이터와 함께 움직입니다)*
* **Excel 2016 / 2019 / 2021:** 함수 이름 뒤에 **`.FLOAT`**만 붙이세요!  
  예: `=SMILES2IMG.FLOAT(A2)`, `=NAME2IMG.FLOAT(A2)` *(셀 테두리에 딱 맞춰 함께 이동하고 크기가 조절되는 그림 도형으로 삽입됩니다)*
:::

---

## 2. 마우스로 직접 그리기 & 화학 반응식 (`Ctrl + Shift + K`)

SMILES 코드를 모르거나 다단계 유기 합성 반응식을 넣고 싶을 때는 내장된 **Ketcher 2D 화학 편집기**를 사용하세요.

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. 구조식을 넣을 셀을 클릭하고 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누릅니다 (또는 상단 **`[SMILES] ➔ [구조식 삽입]`** 클릭).
2. 벤젠 고리, 작용기, 입체 결합(쐐기/점선), 또는 **반응 화살표와 촉매 텍스트($H_2SO_4$, 55 °C)**를 자유롭게 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 셀에 즉시 삽입됩니다!
   * 단일 분자는 `=SMILES2IMG(...)` 수식으로 들어가고,
   * 반응 화살표나 텍스트가 포함된 반응식은 위치와 글자가 100% 보존되는 `=KET2IMG(...)` 수식으로 자동 저장됩니다.
4. **다시 고치려면?** 해당 셀을 **마우스로 더블클릭**하면 그렸던 분자나 반응식이 그대로 다시 열립니다.

---

## 3. 부분구조 하이라이트 & 신약 후보 스크리닝 (SAR)

외부 화학정보학 프로그램 없이 엑셀의 **자동 필터**와 **조건부 서식**만으로 수천 개의 화합물을 즉시 분석할 수 있습니다.

### ① 원하는 골격(작용기)에 색칠하기 (`hl:` 옵션)
세 번째 옵션 인수에 `"hl:찾을구조"`를 적으면 해당 부분구조만 눈에 띄게 강조 표시됩니다:
```text
=SMILES2IMG(A2, "", "hl:c1ccccc1")
```

### ② 3단계 가상 스크리닝 레시피
1. **특정 골격 포함 여부 필터링 (`TRUE` / `FALSE`):**
   ```text
   =SMILES.MATCH(A2, "c1ccccc1")
   ```
2. **기준 물질(`$B$1`)과의 2D 구조 유사도 점수 계산 (`0.0 ~ 1.0`):**
   ```text
   =SMILES.SIMILARITY(A2, $B$1)
   ```
3. **경구 투여 약물성(Lipinski's Rule of 5) 및 물리화학 지표 계산:**
   * 통과 여부(`"PASS"` / `"FAIL"`): `=SMILES2LIPINSKI(A2)`
   * 지용성(LogP) 및 극성 표면적(TPSA): `=SMILES2LOGP(A2)`, `=SMILES2TPSA(A2)`
   * 정밀 질량(Exact Mass, HRMS용): `=SMILES2EXACTMASS(A2)`

> 📐 25종 이상의 전체 엑셀 함수 목록은 **[엑셀 함수 치트시트](/ko/reference/formula)**에서 한눈에 볼 수 있습니다.

---

## 4. 자주 쓰는 스타일 옵션 (흑백 · 다크모드 · 원자번호)

`=SMILES2IMG(셀, [배경색], [옵션])`의 세 번째 인수에 원하는 키워드를 공백이나 쉼표로 조합해 넣을 수 있습니다:

| 내가 원하는 스타일 | 수식 입력 예시 | 설명 |
| :--- | :--- | :--- |
| **학술지/특허용 흑백 선화** | `=SMILES2IMG(A2, "", "bw")` | 원소 색상 없이 깔끔한 검은색 선과 글자로 표시 (ACS 논문 표준) |
| **특정 부분구조 하이라이트** | `=SMILES2IMG(A2, "", "hl:C(=O)O")` | 카복실기(`C(=O)O`) 등 지정한 작용기만 컬러로 강조 표시 |
| **원자 번호 & 입체화학 표시** | `=SMILES2IMG(A2, "", "num stereo")` | NMR/MS 피크 해석을 위한 원자 번호(`num`) 및 *R/S*, *E/Z*(`stereo`) 표시 |
| **다크 모드 (어두운 배경 + 흰색 선)** | `=SMILES2IMG(A2, "#1e1e1e", "white")` | 어두운 셀 배경 위에 선명한 흰색 결합선으로 렌더링 |
| **수소 펼치기 / 회전 / 반전** | `=SMILES2IMG(A2, "", "h rot90 flip")` | 생략된 수소 표시(`h`), 90도 회전(`rot90`), 좌우 반전(`flip`) |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 5. 상단 `[SMILES]` 리본 메뉴 & `.sdf` 파일 연동

수식을 일일이 수정하지 않아도 상단 **`[SMILES]`** 리본 탭에서 클릭 한 번으로 시트 전체를 제어할 수 있습니다:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Office" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### 📂 대용량 `.sdf` 화학 파일 가져오기 / 내보내기 (`[SDF File]`)
* **📥 Import `.sdf` File...:** 여러 화합물이 담긴 `.sdf` 파일을 열면, 엑셀이 새 시트에 **구조식 그림(`=SMILES2IMG`), SMILES, 화학식, 분자량, 그리고 파일 내 모든 속성 태그(`<IC50>`, `<Activity>` 등)**가 정리된 엑셀 표를 자동으로 만들어 줍니다.
* **📤 Export `.sdf` File...:** 시트의 화합물 표를 선택해 표준 `.sdf` 파일로 즉시 저장합니다.

### 🎨 시트 전체 스타일 원클릭 통일 (`[Batch Style]`)
발표 자료나 논문 양식에 맞춰 시트 내 모든 구조식의 디자인을 한 번에 바꿉니다:
* **Standard Color:** 표준 원소 컬러 + 투명 배경 (기본 엑셀 작업용)
* **Publication B&W:** ACS 학술지 및 특허 출원용 흑백 선화
* **White Background:** 원소 컬러 + 불투명 흰색 배경 (색상이 칠해진 셀 위에서도 선명하게 표시)
* **Dark Mode:** 흰색 결합선 + 어두운/투명 배경 (다크 테마 시트 및 슬라이드용)

### 🖼️ Illustrator / ChemDraw용 벡터 복사 (`[Copy SVG]`)
* 선택한 셀의 분자를 화질이 절대 깨지지 않는 **벡터 그래픽(`SVG`)**으로 복사합니다. Adobe Illustrator, Inkscape, 파워포인트 등에 바로 붙여넣어 선 두께나 색상을 자유롭게 편집할 수 있습니다.

---

## 6. 우클릭 컨텍스트 메뉴 (`Chemical Structure`)

셀을 마우스 우클릭하면 나타나는 **`Chemical Structure`** 메뉴에서 자주 쓰는 기능을 바로 실행할 수 있습니다:

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ Edit Structure in Ketcher:** 현재 셀의 분자를 편집기에서 열어 제자리 수정
* **🧬 3D Structure (Mol\*)...:** 3D 단백질 뷰어 실행 (셀에 `1CRN` 같은 4자리 PDB 코드가 적혀 있으면 자동으로 해당 단백질을 불러옵니다!)
* **📋 Copy SMILES / InChI / KET / SVG:** 원하는 화학 포맷이나 벡터 그래픽으로 즉시 복사
* **🔬 Copy Molecule Info (분자 정보 복사 & 상세창):**
  * 물질명, 화학식, 분자량, InChIKey를 개별 복사하거나,
  * **엑셀 가로 행(TSV)**, **마크다운 표**, **줄바꿈 목록** 형태로 한 번에 복사해 다른 표에 바로 붙여넣을 수 있습니다.
  * **[상세 정보 창 열기]**를 누르면 분자량, 정밀 질량, LogP, TPSA, Lipinski 통과 여부를 한눈에 보여주는 요약 창이 열립니다.

::: details 🔍 수식이 #VALUE! 오류를 반환할 때 (진단 창 활용법)
참조한 셀의 SMILES 문법이 잘못되었을 때(예: 괄호나 고리 번호가 안 닫힌 경우) `#VALUE!`가 표시됩니다.  
상단 리본의 **`[SMILES] ➔ [설정 및 진단 (Settings & Diagnostics)]`**을 클릭하면 어떤 셀의 SMILES에서 무슨 문법 오류가 났는지 정확한 원인을 실시간 로그로 확인할 수 있습니다.
:::
