---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "오피스에서 바로 그리는 화학 구조식 & 3D 단백질"
  tagline: "외부 프로그램 없이 Excel, Word, PowerPoint 안에서 2D 분자·반응식을 직접 그리고, 수식으로 분자량·약물성을 자동 계산하며, 3D 단백질을 마우스로 돌려보세요. 100% 무료·오프라인으로 작동합니다."
  actions:
    - theme: brand
      text: 🚀 설치 및 1분 첫걸음
      link: /ko/guide/getting-started
    - theme: alt
      text: 📗 Excel 가이드
      link: /ko/guide/excel
    - theme: alt
      text: 📐 엑셀 함수 치트시트
      link: /ko/reference/formula
    - theme: alt
      text: GitHub
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: "엑셀 셀 안에 구조식 자동 생성"
    details: "=SMILES2IMG(A2) 또는 =NAME2IMG('Aspirin') 수식 하나로 수백 개의 화합물 이름·CAS 번호·SMILES를 셀 크기에 딱 맞는 고화질 구조식으로 변환합니다."
  - icon: ✏️
    title: "더블클릭으로 언제든 직접 그리기"
    details: "셀이나 그림을 더블클릭(또는 Ctrl+Shift+K)하면 내장된 Ketcher 2D 화학 편집기가 즉시 열려, 벤젠 고리와 합성 반응식을 자유롭게 그리고 수정할 수 있습니다."
  - icon: 🔬
    title: "분자량 · 화학식 · 약물성(Lipinski) 자동 계산"
    details: "=SMILES2MW, =SMILES2FORMULA, =SMILES2LOGP, =SMILES2LIPINSKI, =SMILES.MATCH 수식으로 엑셀 시트에서 화합물 물성과 부분구조 포함 여부를 즉시 분석합니다."
  - icon: 🧬
    title: "마우스로 돌려보는 3D 단백질 (Mol*)"
    details: "PDB ID(예: 1CRN, 4HHB)만 입력하면 단백질·핵산 3D 구조를 불러와 슬라이드와 문서 안에서 360° 회전 가능한 3D 모델로 바로 삽입합니다."
  - icon: 🎨
    title: "논문용 흑백 · 다크 모드 · SVG 벡터 원클릭 변환"
    details: "문서 전체의 구조식을 클릭 한 번으로 학술지(ACS) 흑백 선화나 다크 모드용 흰색 선으로 통일하고, Illustrator용 무손실 SVG 벡터나 SDF 파일로 주고받습니다."
  - icon: 🛡️
    title: "100% 오프라인 작동 (연구 기밀 완벽 보호)"
    details: "모든 화학 렌더링과 물성 계산이 내 PC 안에서만 처리됩니다. 외부 서버로 구조식이 절대 전송되지 않아 특허 출원 전 화합물이나 신약 후보 물질도 안심하고 다룰 수 있습니다."
---

## ⚡ 10초 빠른 설치 (Excel · Word · PowerPoint 통합)

**PowerShell** 창을 열고 아래 한 줄을 복사해 붙여넣은 뒤 <kbd>Enter</kbd>를 누르세요 (관리자 권한 불필요):

```powershell
# 정식 안정판 (v1.7.0) 설치 및 자동 업그레이드
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 설치 후 바로 해보세요!
설치가 끝나면 오피스를 다시 켜고 상단의 **`[SMILES]`** 탭을 눌러보거나, 엑셀 빈 셀에 `=NAME2IMG("Aspirin")`을 입력해 보세요. 자세한 따라하기는 **[설치 및 1분 첫걸음](/ko/guide/getting-started)**에서 확인할 수 있습니다.
:::

---

## 🎯 지금 어떤 작업을 하고 싶으신가요?

원하는 작업을 클릭하면 해당 사용법으로 바로 이동합니다:

| 내가 하려는 작업 | 가장 빠른 방법 | 바로가기 |
| :--- | :--- | :---: |
| **시약 이름이나 SMILES로 엑셀 표 만들기** | 셀에 `=NAME2IMG("Caffeine")` 또는 `=SMILES2IMG(A2)` 입력 후 아래로 드래그 | [Excel 가이드 ➔](/ko/guide/excel) |
| **직접 마우스로 분자·반응식 그려 넣기** | Excel · Word · PPT 어디서든 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> 누르기 | [2D 스케처 가이드 ➔](/ko/guide/ketcher) |
| **분자량, 화학식, LogP, Lipinski 계산하기** | `=SMILES2MW(A2)`, `=SMILES2FORMULA(A2)`, `=SMILES2LIPINSKI(A2)` 입력 | [함수 치트시트 ➔](/ko/reference/formula) |
| **워드 논문이나 PPT 발표 자료에 구조식 넣기** | 상단 **`[SMILES]`** 탭에서 구조식 삽입, 다크 모드 투명 배경 및 흑백 일괄 변환 | [Word ➔](/ko/guide/word) · [PPT ➔](/ko/guide/powerpoint) |
| **3D 단백질 구조를 슬라이드에 넣고 돌려보기** | 상단 **`[SMILES] ➔ [3D Structure (Mol*)]`** 클릭 후 PDB 코드(예: `1CRN`) 입력 | [3D 뷰어 가이드 ➔](/ko/guide/molstar) |

---

## 🖼️ 오피스 프로그램별 미리보기

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel 가이드](/ko/guide/excel) | [📘 Microsoft Word 가이드](/ko/guide/word) | [📙 Microsoft PowerPoint 가이드](/ko/guide/powerpoint) |
| :---: | :---: | :---: |
| 셀 안 구조식 자동 생성, 분자량·약물성 수식 계산, 부분구조 하이라이트, `.sdf` 파일 입출력 | 문장 흐름에 맞춘 본문/표 삽입, 논문 투고용 흑백(ACS) 일괄 변환, 분자 물성 표 원클릭 붙여넣기 | 슬라이드 정중앙 원클릭 배치, 다크 테마용 투명 배경, 발표 중 360° 회전하는 3D 단백질 모델 |
