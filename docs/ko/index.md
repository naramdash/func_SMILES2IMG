---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "무료 오픈소스 오피스 화학 구조식 추가 기능"
  tagline: "Microsoft Excel, Word, PowerPoint를 위한 100% 오프라인 무료 오픈소스 화학 추가 기능. 고해상도 2D 화학 구조식 자동 렌더링, 실시간 인셀 수식, 더블 클릭 제자리 수정, 무손실 반응식 KET 캔버스 지원."
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
  - icon: 🧬
    title: Mol* 3D 거대분자 뷰어
    details: 단백질, 핵산, 리간드 복합체를 워드와 파워포인트에서 3D로 시각화합니다. 회전 가능한 GLB 3D 모델 및 논문용 2D 이미지를 삽입하고 더블 클릭으로 재편집할 수 있습니다.
  - icon: 🔬
    title: 원클릭 분자 정보 복사 및 계산 함수
    details: 우클릭 서브메뉴로 분자식, 분자량, 물질명, SMILES, InChI, InChIKey를 플로팅 배지와 함께 즉시 복사합니다. 엑셀에서는 =SMILES2MW, =SMILES2NAME 등의 수식으로 직접 계산합니다.
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

::: tip 💡 간편한 업데이트
나중에 최신 버전이 출시되었을 때도 동일한 명령어를 실행하면 최신 버전으로 자동 업그레이드됩니다!
:::

---

## 🖼️ 오피스 프로그램별 주요 화면

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel 가이드](/ko/guide/excel) | [📘 Microsoft Word 가이드](/ko/guide/word) | [📙 Microsoft PowerPoint 가이드](/ko/guide/powerpoint) |
| :---: | :---: | :---: |
| 네이티브 인셀 수식(`=SMILES2IMG`), 행/열 연동, 정렬/필터 지원, 분자 특성 계산 함수 | 문장 흐름을 따르는 인라인 그림, 표 셀 서식 완벽 보존, Mol* 3D 모델 삽입 | 슬라이드 정중앙 원클릭 배치, 다크 테마 투명 배경 지원, 회전 가능한 3D GLB 모델 |

| [✏️ Ketcher 2D 에디터](/ko/guide/ketcher) | [🧬 Mol* 3D 뷰어](/ko/guide/molstar) | [📐 수식 레퍼런스](/ko/reference/formula) |
| :---: | :---: | :---: |
| 100% 오프라인 분자 및 반응식 스케치, KET 무손실 캔버스 | 생체 고분자 3D 구조, RCSB PDB 직접 다운로드, 회전 가능한 3D GLB | 렌더링 옵션 및 분자 특성 계산 함수 상세 규격 |
