# SMILES for Microsoft PowerPoint 가이드

**SMILES for Microsoft PowerPoint**는 학회 발표, 연구 세미나, 강의 슬라이드에 2D 화학 구조식을 원클릭으로 배치하고, 다크 테마를 위한 투명 배경 및 더블 클릭 제자리 수정을 지원하는 프레젠테이션용 화학 추가 기능입니다.

---

## 🌟 한눈에 보는 핵심 기능

전문 학술 슬라이드 및 학회 발표 자료를 만들 때 위치 어긋남 없이 화학 구조식과 반응식을 깔끔하게 배치할 수 있습니다:

<p align="center">
  <img src="/powerpoint-showcase.png" alt="SMILES for PowerPoint Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip PowerPoint 슬라이드 제작 핵심 장점
* **다크 테마 투명 배경 최적화:** 흰색 사각형 테두리 없이 깔끔한 알파 채널 투명 배경으로 렌더링되어 다크 모드 슬라이드 템플릿에 자연스럽게 녹아듭니다.
* **슬라이드 위치 및 비율 완벽 보존:** 슬라이드 위 분자 그림을 더블 클릭하여 수정해도 기존 좌표(`Left`, `Top`)와 중심축, 크기 비율이 흐트러지지 않고 제자리에서 갱신됩니다.
* **무손실 슬라이드 데이터 보존:** 파워포인트 도형 내부 속성(`Shape.Tags['KET']`와 `Shape.Tags['SMILES']`)에 원본 구조식 정보가 영구 저장되므로, 다른 PC로 PPTX 파일을 공유해도 언제든 재편집할 수 있습니다.
:::

---

## 📑 전용 오피스 리본 메뉴: `[SMILES]` 탭

상단 리본 메뉴의 `[SMILES]` 탭에서 슬라이드 원클릭 삽입, 클립보드 즉시 변환, 시스템 진단을 원스톱으로 제공합니다:

<p align="center">
  <img src="/powerpoint-ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ 구조식 삽입 (`Ctrl+Shift+K`):** 오프라인 Ketcher 2D 에디터에서 구조식이나 화학 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 현재 슬라이드 정중앙에 고해상도 그림으로 자동 배치됩니다.
2. **🧬 3D 구조식 (Mol*):** 내장된 Mol* 3D 뷰어를 실행합니다. RCSB PDB ID(예: `1CRN`, `4HHB`)를 직접 다운로드하거나 로컬 구조 파일을 열어, 슬라이드 쇼 중에도 마우스로 회전/확대할 수 있는 **인터랙티브 3D 모델 (.glb)** 또는 고해상도 2D 이미지로 삽입합니다. 자세한 사용법은 [Mol* 3D 뷰어 가이드](/ko/guide/molstar)를 참고하세요.
3. **⚡ 퀵 인서트 (Quick Insert):**
   - 슬라이드의 텍스트 상자나 불릿 기호 안의 SMILES 텍스트를 선택한 상태에서 누르면 0.1초 만에 분자 도형으로 즉시 치환됩니다.
   - 선택된 개체가 없으면 클립보드의 SMILES를 슬라이드 중앙에 즉시 배치합니다.
4. **📋 구조식 붙여넣기:** 외부 논문이나 ChemDraw 등에서 복사한 화학 표기를 파워포인트 슬라이드 위에 고화질 이미지로 삽입합니다.
5. **🖼️ Copy SVG *(v1.7.0 신규)*:** 선택된 분자 도형을 무손실 벡터 그래픽(`SVG`)으로 렌더링하여 클립보드에 복사합니다 (`image/svg+xml` 스트림 및 raw XML 텍스트 동시 복사로 Adobe Illustrator, Inkscape 등에 무손실 붙여넣기 지원).
6. **🎨 Batch Style *(v1.7.0 신규)*:** 현재 프레젠테이션 전체 슬라이드(또는 선택 영역)의 모든 화학 구조식을 클릭 한 번으로 통일된 스타일 프리셋(**표준 컬러**, **논문/특허용 흑백**, **흰색 배경**, **다크 모드**)으로 일괄 변환합니다.
7. **🔬 Molecule Info 드롭다운:** 발표 슬라이드 글머리 기호(불릿)용 물성 복사, 또는 전체 **Molecule Info 상세 창**을 실행합니다.
8. **🔍 설정 및 진단 (Settings & Diagnostics):** Indigo 엔진 로딩 상태 및 실시간 렌더링 로그를 확인합니다.

---

### 🎨 프레젠테이션 전체 일괄 스타일 변환 (Batch Style)

발표 슬라이드 템플릿(밝은 배경 vs 다크 테마)에 맞춰 슬라이드 덱 내 수십 개의 화학 구조식을 원클릭으로 통일할 수 있습니다:

| 스타일 프리셋 | 리본 아이콘 (`imageMso`) | 슬라이드 시각적 결과 및 권장 용도 |
| :--- | :--- | :--- |
| **Standard Color** | `ColorMenu` | IUPAC CPK 원소 색상 + **투명 배경** (밝은색 배경 및 그라디언트 슬라이드용). |
| **Publication B&W** | `BlackAndWhiteGrayscale` | 고대비 **ACS 학술 논문용 흑백 선화** (세미나 배포용 인쇄 슬라이드 표준). |
| **White Background** | `BlackAndWhiteLightGrayscale` | CPK 원소 색상 + **불투명 흰색 사각형 배경** (사진이나 텍스처 배경 위에서도 구조식을 또렷하게 강조). |
| **Dark Mode** | `BlackAndWhiteBlack` | 고대비 **흰색 결합선 잉크 + 투명 배경** (다크 테마 템플릿 및 OLED 빔프로젝터 발표에 최적화). |

---

## 🖱️ 개편된 우클릭 컨텍스트 메뉴: `Chemical Structure`

슬라이드 위 분자 도형이나 빈 슬라이드 캔버스에서 우클릭을 하면 깔끔한 화학 전용 서브메뉴가 나타납니다:

<p align="center">
  <img src="/powerpoint-context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### A. 슬라이드의 화학 분자 도형 또는 3D 모델을 우클릭한 경우
* **✏️ Edit Structure in Ketcher:** 내장 Ketcher 에디터를 열어 제자리에서 바로 재편집합니다.
* **🧬 Edit 3D Structure (Mol*):** 내장 Mol* 뷰어를 열어 작업하던 각도와 표현 스타일 그대로 재편집합니다.
* **📋 Copy SMILES:** 표준 SMILES 문자열을 클립보드로 복사합니다.
* **📋 Copy InChI:** 표준 IUPAC InChI 문자열을 즉시 변환하여 복사합니다.
* **📋 Copy Canvas Data (KET):** 반응식이 포함된 KET JSON 데이터를 복사합니다.
* **🖼️ Copy Vector Graphic (SVG):** 분자 도형을 무손실 벡터 그래픽(`SVG`)으로 렌더링하여 클립보드에 복사합니다.
* **🔬 Copy Molecule Info (분자 정보 서브메뉴 & 상세창):**
  * **물질명 (Substance Name):** PubChem/IUPAC 기반 일반 물질명 복사 (`Aspirin`, `Ethanol` 등).
  * **화학식 (Chemical Formula):** Hill 표기법 화학식 복사 (`C9H8O4` 등).
  * **분자량 (Molecular Weight):** 정밀 분자량 복사 (`180.16 g/mol` 등).
  * **SMILES / InChI / InChIKey:** 각 표준 식별자 원클릭 즉시 복사.
  * **전체 요약 복사 (Copy All Summary):** 모든 정보를 한 줄 텍스트로 일괄 복사.
  * **여러 줄 리스트로 복사 (Copy as Multi-Line List):** 슬라이드 텍스트 상자에 붙여넣으면 즉시 글머리 기호(Bullet points) 목록으로 들어갑니다!
  * **표 행(TSV Columns) / 마크다운 표 복사:** 슬라이드 표나 엑셀에 바로 붙여넣을 수 있는 표 행 서식 복사.
  * **상세 정보 창 열기...:** 물성 행별 **GDI+ 벡터 케미컬 아이콘**, DataGridView 표, **LogP / TPSA**, **Lipinski Rule of 5**, 다중 서식 내보내기 Split Button(JSON 포함)이 탑재된 전문 분자 물성 다이얼로그 실행.
  > ⚡ **모달 창 방해 없는 플로팅 토스트:** 복사 항목을 누르면 팝업 확인 창이 뜨는 대신 마우스 커서 바로 옆에 가벼운 플로팅 배지(`✓ Copied {항목}: {값}`)가 1.4초간 나타났다 사라지며, 발표 준비 중 작업 포커스를 뺏지 않습니다!

### B. 빈 슬라이드 배경을 우클릭한 경우
* **✏️ Insert Structure in Ketcher:** Ketcher 에디터를 열어 새 구조식을 그린 후 슬라이드에 삽입합니다.
* **🧬 3D Structure (Mol*)...:** Mol* 뷰어를 열어 3D 고분자 모델이나 이미지를 슬라이드에 삽입합니다.
* **📋 Paste Structure:** 클립보드의 화학식 텍스트를 고화질 슬라이드 도형으로 즉시 배치합니다.

---

## 🧪 내장 Ketcher 2D 화학 에디터

슬라이드 위 화학 도형을 더블 클릭하거나 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르면 오프라인 Ketcher 에디터가 실행됩니다 (WebView2 웜 인스턴스 풀링으로 첫 실행 이후 `<0.05초` 만에 즉시 열림):

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **🌓 다크 테마 투명 배경 전환 (`Background: Transparent`):** Ketcher 상단 상태 바에서 투명 배경을 활성화하면 알파 채널이 보존된 고해상도 PNG로 내보내어 어두운 학회 발표 템플릿에서도 완벽하게 어우러집니다.
* **상단 실시간 물성 계산:** 그리는 도중에도 화학식(Formula), 분자량(MW), 정밀 질량(Exact Mass)이 상단에 실시간으로 집계됩니다.
* **단축키:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd>로 슬라이드에 즉시 반영, <kbd>Esc</kbd>로 취소.

---

## 🔬 다단계 반응 경로 & 유기 합성 스킴

학술 세미나에서 다단계 반응 메커니즘, 촉매 순환 주기, 합성 경로를 슬라이드에 표현할 때도 완벽하게 지원됩니다:

* **반응식 무손실 보존:** 반응 화살표, 촉매 조건($H_2SO_4$, 55 °C), 중간체 및 부산물 구조가 파워포인트 도형 메타데이터(`Shape.Tags['KET']`)에 온전히 유지됩니다.
* **슬라이드 제자리 수정:** 발표 직전 반응물이나 시약을 수정해야 할 때, 슬라이드에서 더블 클릭하면 Ketcher에 전체 반응 경로가 그대로 열려 즉시 수정할 수 있습니다.

---

## 🔍 내장 진단 콘솔 (Diagnostics)

리본 메뉴의 **[SMILES] ➔ [설정 및 진단]**을 클릭하여 추가 기능 구동 상태를 점검할 수 있습니다:

* PowerPoint COM Add-in 등록 상태 및 C++ Indigo 화학 엔진 상태를 확인합니다.
* WebView2 런타임 및 오프라인 Ketcher 번들 파일 무결성을 점검합니다.
