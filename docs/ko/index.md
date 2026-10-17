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
      text: 🔬 물성 · SAR 스크리닝
      link: /ko/guide/excel-cheminformatics
    - theme: alt
      text: 🛠️ 엑셀 함수 표현 생성기
      link: /ko/tools/formula-builder

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

---

## 🖥️ 오피스 프로그램별 핵심 화면 (탭을 클릭해 보세요!)

<OfficeAppTabs />

---

## 🎯 지금 필요한 기능을 바로 찾아보세요

| 내가 하려는 작업 | 가장 빠른 방법 | 상세 가이드 |
| :--- | :--- | :---: |
| **처음 설치하고 첫 분자 그림 띄워보기** | PowerShell 1줄 설치 후 `=NAME2IMG("Aspirin")` 또는 <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>K</kbd> | [설치 및 1분 첫걸음 ➔](/ko/guide/getting-started) |
| **시약 이름이나 SMILES로 엑셀 구조식 표 만들기** | `=SMILES2IMG(A2)`, `=NAME2IMG("50-78-2")` 입력 후 아래로 드래그 | [Excel 기초 가이드 ➔](/ko/guide/excel) |
| **분자량, 정밀질량(HRMS), Lipinski, SAR 유사도 분석** | `=SMILES2MW`, `=SMILES2LIPINSKI`, `=SMILES.MATCH`, `.sdf` 입출력 | [Excel 심화 (SAR·SDF) ➔](/ko/guide/excel-cheminformatics) |
| **워드 논문이나 PPT 슬라이드에 구조식 넣기** | 문단 인라인 삽입(Word), 다크 모드 투명 배경 & 3D 회전 발표(PPT) | [Word ➔](/ko/guide/word) · [PowerPoint ➔](/ko/guide/powerpoint) |
| **논문 투고용 흑백 통일 · Illustrator SVG · 물성 표 복사** | 리본 `[Batch Style]`, `[Copy SVG]`, `[Molecule Info]` 다중 서식 내보내기 | [출판 스타일 · SVG · 물성 ➔](/ko/guide/publishing) |
| **마우스로 직접 2D 반응식 그리거나 3D 단백질 넣기** | Ketcher 2D 스케처(`=KET2IMG`) 및 Mol* 3D 단백질 뷰어(`.glb`) | [2D 스케처 ➔](/ko/guide/ketcher) · [3D 단백질 ➔](/ko/guide/molstar) |
| **25종 이상의 전체 엑셀 함수 목록과 인수 문법 찾기** | 목적별로 정리된 함수 사전 및 스타일 키워드 표 확인 | [전체 함수 치트시트 ➔](/ko/reference/formula) |
| **마우스 클릭만으로 엑셀 수식과 스타일 조합해 복사하기** | 웹 수식 생성기 위젯에서 옵션 체크 후 실시간 미리보기 & 원클릭 복사 | [🛠️ 함수 표현 생성기 ➔](/ko/tools/formula-builder) |
