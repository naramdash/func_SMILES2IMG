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
  <img src="/ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ 구조식 삽입 (`Ctrl+Shift+K`):** 오프라인 Ketcher 에디터에서 구조식이나 화학 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>(또는 적용 버튼)를 누르면 현재 슬라이드 정중앙에 고해상도 그림으로 자동 배치됩니다.
2. **⚡ 퀵 인서트 (Quick Insert):**
   - 슬라이드의 텍스트 상자나 불릿 기호 안의 SMILES 텍스트를 선택한 상태에서 누르면 0.1초 만에 분자 도형으로 즉시 치환됩니다.
   - 선택된 개체가 없으면 클립보드의 SMILES를 슬라이드 중앙에 즉시 배치합니다.
3. **📋 클립보드에서 붙여넣기:** 외부 논문이나 ChemDraw 등에서 복사한 화학 표기를 파워포인트 슬라이드 위에 고화질 이미지로 삽입합니다.
4. **🔍 설정 및 진단 (Settings & Diagnostics):** Indigo 엔진 로딩 상태 및 실시간 렌더링 로그를 확인합니다.

---

## 🖱️ 개편된 우클릭 컨텍스트 메뉴: `Chemical Structure`

슬라이드 위 분자 도형이나 빈 슬라이드 캔버스에서 우클릭을 하면 깔끔한 화학 전용 서브메뉴가 나타납니다:

<p align="center">
  <img src="/context-menu-guide.png" alt="Chemical Structure Context Menu in PowerPoint" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### A. 슬라이드의 화학 분자 도형을 우클릭한 경우
* **✏️ Edit Structure in Ketcher:** 내장 에디터를 열어 슬라이드 위 분자를 제자리에서 바로 재편집합니다.
* **📋 Copy SMILES / Copy Molfile (KET):** 화학 구조식 텍스트를 클립보드로 복사합니다.
* **🔬 Copy Molecular Formula:** 표준 Hill 표기법 화학식(예: `C8H10N4O2`)을 복사합니다.
* **⚖️ Copy Molecular Weight (MW):** 발표 질의응답 시 즉시 활용할 수 있는 IUPAC 정밀 분자량(g/mol)을 복사합니다.
* **🎯 Copy Exact Mass:** 고해상도 단일 동위원소 질량(Monoisotopic Exact Mass)을 복사합니다.

### B. 빈 슬라이드 배경을 우클릭한 경우
* **✏️ Insert Structure in Ketcher:** Ketcher 에디터를 열어 새 구조식을 그린 후 슬라이드에 삽입합니다.
* **📋 Paste from Clipboard:** 클립보드의 화학식 텍스트를 고화질 슬라이드 도형으로 즉시 배치합니다.

---

## 🧪 내장 Ketcher 2D 화학 에디터

슬라이드 위 화학 도형을 더블 클릭하거나 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르면 오프라인 Ketcher 에디터가 실행됩니다:

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
