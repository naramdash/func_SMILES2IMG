# SMILES for Microsoft Word 가이드

**SMILES for Microsoft Word**는 워드 문서, 논문, 연구 보고서 본문 및 표 안에 고해상도 2D 화학 구조식을 인라인 그림으로 삽입하고, 더블 클릭으로 표 서식 붕괴 없이 제자리 수정을 지원하는 추가 기능입니다.

---

## 🎥 3분 핵심 기능 마스터 영상

> 💡 **SMILES for Word의 핵심 기능을 1분 54초 만에 마스터할 수 있는 1080p 고해상도 가이드 영상입니다.**  
> 상단 리본 메뉴 활용, 드래그 텍스트 0.1초 즉시 치환(Quick Insert), 표 레이아웃 유지 더블 클릭 제자리 수정, 무손실 KET 반응식 보존, 우클릭 분자량 계산까지 단계별로 확인하세요.

<p align="center">
  <video src="https://github.com/user-attachments/assets/e34fc58a-b320-480e-8e83-876356796b6e" controls="controls" width="100%" style="max-width: 800px; border-radius: 8px;">
    <a href="https://github.com/user-attachments/assets/e34fc58a-b320-480e-8e83-876356796b6e">🎬 Word 핵심 기능 3분 마스터 동영상 재생하기</a>
  </video>
</p>

<details>
<summary><b>⏱️ 영상 타임라인 및 챕터별 내용</b></summary>

| 시간 | 챕터 및 주제 | 핵심 내용 및 자막 요약 |
| :---: | :--- | :--- |
| **00:00** | **개요 및 핵심 소개** | 고해상도 인라인 그림, 더블 클릭 제자리 수정, 0.1초 즉시 치환 및 무손실 KET 특징 |
| **00:13** | **1단계: 구조식 삽입** | 상단 리본 `[SMILES]` 탭의 `[구조식 삽입]`으로 Ketcher 에디터 실행 및 문서 삽입 |
| **00:28** | **2단계: 퀵 인서트** | 본문 SMILES 텍스트 드래그 또는 클립보드 복사 후 `[퀵 인서트]`로 0.1초 만에 그림 변환 |
| **00:43** | **3단계: 더블 클릭 수정** | 표 서식을 유지한 채 그림을 마우스로 더블 클릭하여 제자리에서 바로 재편집 |
| **00:59** | **4단계: 무손실 반응식** | 반응 화살표와 텍스트가 `AlternativeText`에 영구 보존되어 `.docx` 저장 후에도 무결성 유지 |
| **01:14** | **5단계: 물성 정보 복사** | 우클릭 메뉴로 SMILES 역추출 및 화학식·분자량(MW, g/mol) 즉시 산출 |
| **01:28** | **핵심 요약 & 체크리스트** | 학술 논문 및 연구 보고서 작성을 위한 4대 핵심 체크포인트 정리 |

</details>

---

## 🌟 한눈에 보는 핵심 기능

<p align="center">
  <img src="/word-showcase.png" alt="SMILES for Word Showcase" width="100%" />
</p>

::: tip Word 문서 작업 시 핵심 장점
* **자연스러운 인라인 그림(InlineShape):** 텍스트 흐름에 맞춰 삽입되어 글자처럼 줄바꿈되고, 표 셀 안에 깔끔하게 정렬됩니다.
* **표 레이아웃 보존:** 더블 클릭으로 구조식을 수정해도 표의 열 너비나 행 높이가 깨지지 않고 정확히 그 자리에서 치환됩니다.
* **무손실 영구 보존:** 구조식 원본(`KET` 또는 `SMILES`)이 `.docx` 파일 내부 속성(`AlternativeText`)에 저장되므로, 다른 컴퓨터로 문서를 전송해도 화학 정보가 영구 보존됩니다.
:::

---

## 📑 리본 메뉴: `[SMILES]` 탭

<p align="center">
  <img src="/ribbon-guide.png" alt="Word Ribbon Menu" width="100%" />
</p>

1. **구조식 삽입 (Insert Structure):** 오프라인 Ketcher 2D 에디터를 열고 분자식이나 반응식을 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (또는 적용 버튼)를 누르면 커서 위치에 즉시 삽입됩니다.
2. **퀵 인서트 (Quick Insert):**
   - 본문에서 SMILES 텍스트를 드래그한 상태로 누르면 해당 텍스트가 0.1초 만에 분자 그림으로 즉시 치환됩니다.
   - 드래그한 텍스트가 없으면 클립보드의 SMILES를 읽어 커서 위치에 바로 그림으로 넣습니다.
3. **클립보드에서 붙여넣기:** 클립보드에 복사된 SMILES나 KET를 그림으로 즉시 삽입합니다.
4. **진단 및 로그 (Show diagnostics):** 엔진 로딩 상태, 오류 로그 및 버전 정보를 확인합니다.

---

## 🖱️ 마우스 더블 클릭 제자리 수정

1. 워드 문서나 표 안에 들어있는 분자 구조식 그림을 마우스로 더블 클릭합니다.
2. 내장 Ketcher 에디터가 열리면서 원본 구조식과 결합들이 그대로 로드됩니다.
3. 치환기나 반응식을 수정합니다.
4. **[문서에 적용 (Ctrl+Enter)]**을 누르면 표 너비와 정렬을 유지한 채 그림이 그 자리에서 즉시 업데이트됩니다.

---

## 📋 우클릭 컨텍스트 메뉴

워드 문서 안의 분자 그림이나 빈 영역을 마우스 우클릭:

* **Chemical Structure ➔ Ketcher로 구조 편집:** 오프라인 에디터 실행.
* **Chemical Structure ➔ SMILES / KET 복사:** 순수 화학식 텍스트 추출.
* **Chemical Structure ➔ 분자식(Formula) 복사:** 화학식 계산 (예: `C8H10N4O2`).
* **Chemical Structure ➔ 분자량(MW) 복사:** 분자량 계산 (g/mol, 예: `194.19`).
