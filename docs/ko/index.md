---
layout: home

hero:
  name: "SMILES for Office"
  text: "화학 분자 구조식 자동 렌더러 & 2D 에디터"
  tagline: "고해상도 2D 화학 구조식 자동 렌더링, 실시간 인셀 수식, 더블 클릭 제자리 수정, 무손실 반응식 KET 캔버스 — Microsoft Excel, Word, PowerPoint 100% 로컬 오프라인 구동."
  image:
    src: /excel-showcase.png
    alt: SMILES for Office
  actions:
    - theme: brand
      text: 시작 가이드
      link: /ko/guide/getting-started
    - theme: alt
      text: Excel 가이드
      link: /ko/guide/excel
    - theme: alt
      text: GitHub 저장소
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: 실시간 인셀 수식
    details: 셀 안에 =SMILES2IMG(A2) 입력만으로 고해상도 분자 구조식을 네이티브 삽입하고 행/열 크기 변경 시 자동 연동됩니다.
  - icon: ✏️
    title: 더블 클릭 제자리 수정
    details: 엑셀 셀이나 워드/파워포인트 그림을 마우스로 더블 클릭하면 오프라인 Ketcher 에디터가 즉시 열려 원본 좌표를 유지하며 수정됩니다.
  - icon: 🧪
    title: 무손실 KET 반응식 캔버스
    details: 반응 화살표, 플러스 기호, 화학 텍스트 주석이 문서 파일(.xlsx, .docx, .pptx)에 영구 보존되어 무손실 재편집을 지원합니다.
  - icon: 🛡️
    title: 100% 로컬 오프라인 보안
    details: EPAM Indigo 및 Ketcher v3.18.0이 로컬 바이너리로 내장되어 외부 서버 전송이 전혀 없는 완벽한 보안을 제공합니다.
  - icon: 🎨
    title: 논문 및 특허 출원 표준
    details: ACS 1996 출판 표준 규격, 특허청 흑백 선화(bw), 다크 테마 및 투명 배경 렌더링을 자유롭게 지정할 수 있습니다.
  - icon: 📦
    title: 원클릭 통합 설치기
    details: 단 한 줄의 PowerShell 명령어로 엑셀 XLL, 워드 COM DLL, 파워포인트 COM DLL을 자동 다운로드 및 레지스트리 등록합니다.
---

## ⚡ 빠른 설치 (Quick Install)

PowerShell을 열고 아래 명령어를 실행하면 오피스 전체에 원클릭으로 자동 설치됩니다:

```powershell
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

> **시스템 요구사항:** Windows 10/11, Microsoft Office (Excel / Word / PowerPoint 2016–365, 32비트/64비트 호환), WebView2 Runtime.
