# SMILES for Microsoft Excel 가이드

**SMILES for Microsoft Excel**은 엑셀 워크시트 셀 안에 2D 화학 구조식을 네이티브 인셀 이미지 또는 플로팅 도형으로 실시간 렌더링하고, 편리한 리본 메뉴와 컨텍스트 메뉴, 셀 더블 클릭으로 제자리 수정을 지원하는 화학 연구용 고성능 추가 기능입니다.

---

## 🌟 한눈에 보는 핵심 기능

엑셀에서의 모든 화학 구조식 작업은 상단 리본 메뉴, 우클릭 컨텍스트 메뉴, 그리고 워크시트 수식 엔진과 완벽하게 통합되어 있습니다:

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip ✨ 4줄 요약: 이것만 알면 끝!
1. **빈 셀에 바로 그리기:** <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> 누르고 Ketcher에서 그린 후 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (셀에 즉시 삽입).
2. **수식으로 수백 개 일괄 삽입:** `=SMILES2IMG(A2)` 또는 `=KET2IMG(A2)` 입력 후 아래로 드래그 채우기.
3. **클립보드 즉시 삽입:** SMILES/KET 복사 후 리본 메뉴의 **[클립보드에서 붙여넣기]** 클릭.
4. **언제든 수정 / 물성 확인:** 셀 더블 클릭으로 제자리 수정 / 우클릭 **[Chemical Structure]**로 분자량 및 화학식 원클릭 복사.
:::

---

## 📑 전용 오피스 리본 메뉴: `[SMILES]` 탭

수식 문법이나 단축키를 외우지 않아도 상단 리본 메뉴의 **[SMILES]** 탭에서 모든 화학 기능을 원클릭으로 조작할 수 있습니다:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Office" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **✏️ 구조식 삽입 (`Ctrl+Shift+K`):** 100% 로컬 오프라인 구동되는 Ketcher 2D 화학 에디터를 실행합니다. 구조식이나 화학 반응식을 자유롭게 그린 후 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 현재 선택된 셀에 즉시 삽입됩니다.
* **🧬 3D 구조식 (Mol\*):** 공식 Mol\* 3D 고분자/단백질 뷰어를 실행합니다. RCSB PDB ID로 생체 고분자 구조를 직접 다운로드하거나 로컬 구조 파일을 열어, 회전 가능한 인터랙티브 3D 모델(`.glb`) 또는 고해상도 2D 이미지(`.png`)를 시트에 삽입합니다.
* **⚡ 퀵 인서트 (Quick Insert):** 워크시트에서 SMILES 텍스트가 입력된 셀을 선택한 상태에서 클릭하면 0.1초 만에 인셀 구조식 그림으로 즉시 변환합니다. 빈 셀인 경우 클립보드 내용을 자동 감지하여 삽입합니다.
* **📋 클립보드에서 붙여넣기:** 웹사이트나 논문, ChemDraw 등에서 복사한 SMILES 또는 KET 문자열을 해석하여 현재 셀에 고화질 이미지로 삽입합니다.
* **🖼️ Copy SVG *(v1.7.0 신규)*:** 선택된 분자를 무손실 벡터 그래픽(`SVG`)으로 렌더링하여 클립보드에 복사합니다 (Adobe Illustrator, Inkscape 등에 즉시 붙여넣기 가능).
* **🎨 Batch Style *(v1.7.0 신규)*:** 현재 워크시트 내의 모든 화학 구조식을 클릭 한 번으로 통일된 스타일 프리셋(**투명 컬러**, **논문/특허용 흑백(ACS Monochrome)**, **흰색 배경 컬러**, **다크 모드 흰색 잉크**)으로 일괄 변환합니다.
* **📂 SDF File (`Import .sdf` / `Export .sdf`) *(v1.7.0 신규)*:** 다중 분자가 담긴 표준 `.sdf` 화학 데이터 파일을 읽어 `=SMILES2IMG(...)` 수식과 분자량·화학식·속성 컬럼이 포함된 엑셀 표로 자동 생성하거나, 시트의 분자 목록을 `.sdf` 파일로 내보냅니다.
* **🔍 설정 및 진단 (Settings & Diagnostics):** EPAM Indigo 화학 엔진 구동 상태, 내장 Ketcher 캐시 파일 무결성, 렌더링 벤치마크 로그를 실시간 확인할 수 있는 추가 기능 진단 콘솔 창을 엽니다.

---

## 🖱️ 개편된 우클릭 컨텍스트 메뉴: `Chemical Structure`

엑셀 워크시트 어디에서든 마우스 우클릭을 하면, 엑셀 기본 붙여넣기 옵션과 명확하게 분리된 `Chemical Structure` 서브메뉴를 통해 스마트한 화학 기능에 접근할 수 있습니다:

<p align="center">
  <img src="/excel-context-menu-guide.png" alt="Chemical Structure Context Menu in Excel" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### A. 화학 구조식 셀을 우클릭한 경우
화학 구조식 수식이나 이미지가 들어있는 셀을 우클릭하면 전문 분자 분석 서브메뉴가 나타납니다:
* **✏️ Edit Structure in Ketcher:** 내장 에디터를 열어 해당 분자를 제자리에서 바로 수정합니다.
* **🧬 3D Structure (Mol\*)...:** Mol\* 3D 뷰어를 엽니다. 활성 셀에 4자리 PDB ID(예: `1CRN`)가 적혀 있다면 해당 구조를 자동으로 즉시 불러옵니다!
* **📋 Copy SMILES:** 표준 SMILES 문자열을 클립보드로 복사합니다.
* **📋 Copy InChI:** 표준 IUPAC InChI 문자열을 즉시 변환하여 복사합니다.
* **📋 Copy Canvas Data (KET):** 반응식이 포함된 KET JSON 데이터를 복사합니다.
* **🔬 Copy Molecule Info (분자 정보 서브메뉴):**
  * **물질명 (Substance Name):** PubChem/IUPAC 기반 일반 물질명 복사 (`Aspirin`, `Ethanol` 등).
  * **화학식 (Chemical Formula):** Hill 표기법 화학식 복사 (`C9H8O4` 등).
  * **분자량 (Molecular Weight):** 정밀 분자량 복사 (`180.16 g/mol` 등).
  * **SMILES / InChI / InChIKey:** 각 표준 식별자 원클릭 즉시 복사.
  * **전체 요약 복사 (All Summary):** 모든 정보를 한 줄 텍스트로 일괄 복사.
  * **상세 정보 창 열기...:** 2D SVG 그래픽 뷰어 및 전체 속성 표가 포함된 상세 다이얼로그 창 실행.
  > ⚡ **모달 창 방해 없는 플로팅 토스트:** 복사 항목을 누르면 팝업 확인 창이 뜨는 대신 마우스 커서 바로 옆에 가벼운 플로팅 배지(`✓ Copied {항목}: {값}`)가 1.4초간 나타났다 사라지며, 작업 포커스를 뺏지 않습니다!

### B. 빈 셀이나 구조식 도형을 우클릭한 경우
* **✏️ Insert Structure in Ketcher:** Ketcher 에디터를 실행하여 새 화학 구조식을 그린 후 셀에 즉시 삽입합니다.
* **🧬 3D Structure (Mol\*)...:** Mol\*을 열어 선택된 셀 위치에 3D 모델이나 2D 이미지를 삽입합니다.
* **✏️ Edit 3D Structure (Mol\*)...:** (Mol\* 3D 모델 또는 이미지를 우클릭한 경우) 기존 카메라 각도와 스타일 그대로 Mol\*에서 다시 열어 수정합니다.
* **📋 Paste Structure:** 클립보드에 복사된 화학식 텍스트를 감지하여 렌더링된 구조식으로 즉시 삽입합니다.

---

## 🧪 내장 Ketcher 2D 화학 에디터

화학 셀을 더블 클릭하거나 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르면 오프라인으로 번들링된 전문 Ketcher 2D 화학 에디터 창이 나타납니다:

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### Ketcher 에디터 주요 편의 기능
* **상단 실시간 물성 정보 배너:** 결합이나 원자를 추가/수정하는 즉시 상단 상태 표시줄에 **Formula(화학식)**, **MW(분자량)**, **Exact Mass(정밀 질량)**이 실시간으로 자동 계산되어 표시됩니다.
* **투명 배경 토글 버튼 (`Background: White` ⇄ `Background: Transparent`):** 클릭 한 번으로 배경 투명화 모드를 전환할 수 있습니다. 투명 배경 모드에서는 알파 채널 PNG가 생성되어 다크 모드 엑셀 테마나 발표 슬라이드에서도 사각형 흰 테두리 없이 깔끔하게 표시됩니다.
* **풍부한 화학 도구 지원:** 주기율표 원소 선택기, 벤젠/시클로펜탄/시클로헥산 등 자주 쓰이는 고리 템플릿, 쐐기/점선 입체 결합, 다단계 화학 반응 화살표 및 반응 조건 텍스트 도구를 완벽 지원합니다.
* **초고속 단축키:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 엑셀 워크시트에 즉시 반영되며, <kbd>Esc</kbd>를 누르면 변경을 취소하고 창을 닫습니다.

---

## 🔬 무손실 화학 반응식 캔버스: `=KET2IMG`

단일 분자뿐 아니라 반응 화살표, 촉매($H_2SO_4$), 반응 온도(55 °C), 시간, 플러스 기호(`+`)가 포함된 전체 유기 합성 반응식을 셀 안에 통째로 보존하려면 `=KET2IMG` 수식을 사용하세요:

```text
=KET2IMG(A2)
```

* **무손실 좌표 및 텍스트 보존:** 분자 연결 정보만 담는 SMILES와 달리, Ketcher 고유의 `KET` 포맷은 반응 화살표, 촉매 조건 텍스트, 2D 공간 좌표 레이아웃을 100% 무손실 상태로 수식 안에 영구 보존합니다.
* **더블 클릭 제자리 수정:** 반응식 셀을 더블 클릭하면 Ketcher에 반응 화살표와 텍스트 주석이 그대로 다시 열려 언제든 조건이나 반응물을 제자리에서 수정할 수 있습니다.

---

## 📝 수식 함수: 네이티브 인셀 vs. 플로팅 이미지

엑셀 버전에 따라 최적의 수식 함수를 선택할 수 있습니다:

### 1. `=SMILES2IMG(smiles_or_ket, [bg], [options])` 및 `=INCHI2IMG(inchi, [bg], [options])`
* **지원 버전:** **Microsoft 365** 및 **Excel 2024 이상**
* SMILES, Ketcher KET 또는 IUPAC InChI로부터 셀 내부에 진짜 **인셀(In-Cell) 그림**으로 쏙 들어갑니다.
* 행 높이와 열 너비를 조절하면 분자 그림이 비율을 유지하며 자동으로 크기가 맞춰집니다.
* 정렬, 필터, 셀 복사/이동 시 셀 데이터와 일체형으로 함께 동작합니다.
* VLOOKUP, INDEX/MATCH 등 표준 엑셀 함수 및 조건부 서식과 완벽하게 호환됩니다.

### 2. `=SMILES2IMG.FLOAT(smiles_or_ket, [bg], [options])` 및 `=INCHI2IMG.FLOAT(inchi, [bg], [options])`
* **지원 버전:** **Excel 2016 / 2019 / 2021** (또는 플로팅 도형으로 배치가 필요한 경우)
* 셀 좌표에 고정(`xlMoveAndSize`)되는 플로팅 도형(Shape)으로 삽입됩니다.
* 수식이 변경되면 고아 도형이 남지 않고 자동으로 교체되며, 셀 이동 및 크기 조절에 연동됩니다.

### 3. 분자 화학 속성 계산 함수군
워크시트 수식을 통해 분자 특성을 자동으로 산출할 수 있습니다:
* `=SMILES2NAME(cell)` / `=KET2NAME(cell)` / `=INCHI2NAME(cell)` — PubChem/IUPAC 기반 물질 일반명 자동 조회.
* `=SMILES2MW(cell)` / `=KET2MW(cell)` / `=INCHI2MW(cell)` — 표준 분자량(MW, g/mol) 자동 계산 (0ms 순수 로컬).
* `=SMILES2FORMULA(cell)` / `=KET2FORMULA(cell)` / `=INCHI2FORMULA(cell)` — Hill 표기법 화학식 산출 (`C9H8O4` 등, 0ms 순수 로컬).
* `=SMILES2INCHI(cell)` / `=KET2INCHI(cell)` — IUPAC 표준 InChI 문자열 변환 (0ms 순수 로컬).
* `=INCHI2SMILES(cell)` — IUPAC InChI를 Canonical SMILES로 변환 (0ms 순수 로컬).
* `=SMILES2INCHIKEY(cell)` / `=KET2INCHIKEY(cell)` / `=INCHI2INCHIKEY(cell)` — 27자리 InChIKey 고유 해시 생성 (0ms 순수 로컬).

---

## 🎨 화학 스타일링 & 렌더링 옵션

옵션 키워드는 공백, 쉼표, 세미콜론, 파이프 기호(`|`)로 자유롭게 조합하여 전달할 수 있습니다:

| 옵션 키워드 | 상세 설명 | 사용 예시 |
| :--- | :--- | :--- |
| `bw` / `monochrome` | 특허청(USPTO) 및 ACS 논문 제출용 고대비 흑백 선화 모드 | `=SMILES2IMG(A2, "", "bw")` |
| `colored` | 표준 IUPAC 원소별 고유 컬러 표시 모드 | `=SMILES2IMG(A2, "", "colored")` |
| `num` / `numbers` | 원자 인덱스 번호 강제 표시 (NMR 피크 및 MS 프래그먼트 지정 시 유용) | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | 카이랄 중심 CIP 입체화학 레이블(*R/S*, *E/Z*) 강제 계산 및 표시 | `=SMILES2IMG(A2, "", "stereo")` |
| `d=2.0` / `d=1.5` | 고해상도 수퍼샘플링 배율 (초고선명 인쇄 및 래스터화) | `=SMILES2IMG(A2, "", "d=2.0")` |
| `flip` / `flipv` | 분자 구조 좌우(수평) 또는 상하(수직) 대칭 반전 | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` | 시계 방향 90도 단위 분자 회전 | `=SMILES2IMG(A2, "", "rot90")` |
| `w=300 h=200` | 플로팅 이미지용 픽셀 가로/세로 크기 직접 지정 | `=SMILES2IMG.FLOAT(A2, "", "w=300 h=200")` |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## 🔍 내장 진단 창 및 화학 엔진 실시간 로그

수식에 에러가 나거나 추가 기능의 구동 상태를 점검하고 싶을 때 리본 메뉴의 **[SMILES] ➔ [설정 및 진단 (Settings & Diagnostics)]**을 클릭하세요:

* **EPAM Indigo 엔진 로딩 상태:** C++ 네이티브 화학 엔진 초기화 성공 여부 및 버전 정보를 확인합니다.
* **WebView2 및 Ketcher 캐시 검증:** 로컬에 캐시된 69개 에디터 파일의 SHA256 해시 무결성을 실시간 검증합니다.
* **화학식 구문 파싱 오류 실시간 추적:** 닫히지 않은 고리(Unclosed cycle)나 유효하지 않은 원자가(Valency) 등 비정상적인 SMILES 입력 시, 단순 `#VALUE!` 표시에 그치지 않고 Indigo 엔진의 정확한 파싱 실패 원인을 로그로 알려줍니다.
* **렌더링 성능 벤치마크:** 분자 1건당 렌더링 소요 시간(보통 1.5~3.0 ms)을 실시간 측정하여 표시합니다.
