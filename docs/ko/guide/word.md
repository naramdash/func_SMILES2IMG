# SMILES for Microsoft Word 가이드

**SMILES for Microsoft Word**는 워드 문서, 논문, 연구 보고서 본문 및 표 안에 고해상도 2D 화학 구조식을 인라인 그림으로 삽입하고, 더블 클릭으로 표 서식 붕괴 없이 제자리 수정을 지원하는 화학 학술 연구용 추가 기능입니다.

---

## 🌟 한눈에 보는 핵심 기능

워드에서의 모든 화학 구조식 작업은 문장의 텍스트 흐름에 자연스럽게 녹아들며, 표 셀 안에서도 완벽한 정렬과 제자리 수정을 지원합니다:

<p align="center">
  <img src="/word-showcase.png" alt="SMILES for Word Showcase" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

::: tip Word 문서 작업 시 핵심 장점
* **자연스러운 인라인 그림(InlineShape):** 텍스트 흐름에 맞춰 삽입되어 글자처럼 줄바꿈되고, 표 셀 안에 깔끔하게 정렬됩니다.
* **표 레이아웃 무결성 보존:** 더블 클릭으로 구조식을 수정해도 표의 열 너비나 행 높이가 깨지지 않고 정확히 그 자리에서 치환됩니다.
* **무손실 영구 보존:** 구조식 원본(`KET` 또는 `SMILES`)이 `.docx` 파일 내부 속성(`AlternativeText`)에 저장되므로, 다른 컴퓨터로 문서를 전송해도 화학 정보가 영구 보존됩니다.
:::

---

## 📑 전용 오피스 리본 메뉴: `[SMILES]` 탭

상단 리본 메뉴의 `[SMILES]` 탭에서 화학 구조식 삽입, 0.1초 텍스트 치환, 진단 창을 원클릭으로 조작할 수 있습니다:

<p align="center">
  <img src="/ribbon-guide.png" alt="SMILES Ribbon Tab in Microsoft Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

1. **✏️ 구조식 삽입 (`Ctrl+Shift+K`):** 100% 로컬 오프라인 구동되는 Ketcher 2D 화학 에디터를 열고 분자식이나 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (또는 적용 버튼)를 누르면 커서 위치에 즉시 삽입됩니다.
2. **⚡ 퀵 인서트 (Quick Insert):**
   - 본문에서 SMILES 텍스트를 드래그한 상태로 누르면 해당 텍스트가 0.1초 만에 분자 그림으로 즉시 치환됩니다.
   - 드래그한 텍스트가 없으면 클립보드의 SMILES를 읽어 커서 위치에 바로 그림으로 넣습니다.
3. **📋 클립보드에서 붙여넣기:** 외부 논문이나 웹페이지에서 복사한 SMILES 또는 KET를 읽어 커서 위치에 인라인 그림으로 삽입합니다.
4. **🔍 설정 및 진단 (Settings & Diagnostics):** Indigo 엔진 및 COM 추가 기능 연동 상태와 로그를 확인합니다.

---

## 🖱️ 개편된 우클릭 컨텍스트 메뉴: `Chemical Structure`

워드 본문 또는 표 안에서 마우스 우클릭을 하면 깔끔하게 분리된 `Chemical Structure` 서브메뉴가 나타납니다:

<p align="center">
  <img src="/context-menu-guide.png" alt="Chemical Structure Context Menu in Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

### A. 삽입된 화학 그림을 우클릭한 경우
* **✏️ Edit Structure in Ketcher:** 내장 에디터를 열어 해당 분자를 제자리에서 바로 재편집합니다.
* **📋 Copy SMILES / Copy Molfile (KET):** 표준 화학 표기법 문자열을 클립보드로 복사합니다.
* **🔬 Copy Molecular Formula:** 표준 Hill 표기법 화학식(예: `C8H10N4O2`)을 복사합니다.
* **⚖️ Copy Molecular Weight (MW):** IUPAC 정밀 분자량(g/mol, 예: `194.19`)을 복사합니다.
* **🎯 Copy Exact Mass:** 고해상도 단일 동위원소 질량(예: `194.0804`)을 복사합니다.

### B. 본문 텍스트나 빈 문단을 우클릭한 경우
* **✏️ Insert Structure in Ketcher:** Ketcher 에디터를 열어 새 구조식을 그린 후 커서 위치에 삽입합니다.
* **📋 Paste from Clipboard:** 클립보드의 화학식 텍스트를 인라인 이미지로 즉시 삽입합니다.

---

## 🧪 내장 Ketcher 2D 화학 에디터

워드 본문의 화학 그림을 더블 클릭하거나 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누르면 오프라인 Ketcher 에디터가 실행됩니다:

<p align="center">
  <img src="/ketcher-guide.png" alt="Embedded Ketcher 2D Editor Window" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **실시간 물성 정보 표시:** 원자와 결합을 그리는 즉시 상단 상태 표시줄에 화학식, 분자량, 정밀 질량이 실시간으로 계산됩니다.
* **투명 배경 지원:** 다크 모드 및 컬러 서식 논문 작성을 위해 투명 배경 토글 버튼을 제공합니다.
* **표 셀 자동 맞춤:** 삽입되는 그림은 워드 문서 본문 폭 및 표 셀 크기에 알맞게 자동 스케일링됩니다.
* **단축키:** <kbd>Ctrl</kbd> + <kbd>Enter</kbd>로 워드 문서에 즉시 반영, <kbd>Esc</kbd>로 취소.

---

## 🔬 학술 논문용 화학 반응식 & 합성 경로

화살표, 촉매 조건 텍스트, 플러스 기호(`+`)가 포함된 다단계 유기 합성 경로도 워드 문서에 깔끔하게 삽입할 수 있습니다:

<p align="center">
  <img src="/reaction-scheme-guide.png" alt="Reaction Scheme and Lossless Storage" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* **전체 반응 경로 무손실 보존:** 반응 화살표, 시약($H_2SO_4$), 반응 온도(55 °C) 등이 워드 문서의 도형 속성(`AlternativeText`)에 무손실 저장됩니다.
* **더블 클릭 제자리 수정:** 논문 투고 전 반응 조건이나 수율을 수정해야 할 때, 더블 클릭만으로 전체 반응 스킴이 Ketcher에 열려 손쉽게 수정할 수 있습니다.

---

## 🔍 내장 진단 콘솔 (Diagnostics)

리본 메뉴의 **[SMILES] ➔ [설정 및 진단]**을 클릭하여 추가 기능 구동 상태를 점검할 수 있습니다:

<p align="center">
  <img src="/diagnostics-guide.png" alt="SMILES Diagnostics Console in Word" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

* Word COM Add-in 등록 상태 및 C++ Indigo 화학 엔진 초기화 상태를 확인합니다.
* 로컬 캐시 파일 무결성을 점검하고 실시간 렌더링 소요 시간 벤치마크를 모니터링합니다.
