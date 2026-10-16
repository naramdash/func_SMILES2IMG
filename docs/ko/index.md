---
layout: home

hero:
  name: "func_SMILES2IMG"
  text: "무료 오픈소스 오피스 통합 화학 추가 기능"
  tagline: "Microsoft Excel, Word, PowerPoint를 위한 100% 오프라인 오픈소스 화학 스위트. 2D 구조식 렌더링(SMILES · InChI · KET · 물질명/CAS), Ketcher 2D 에디터, Mol* 3D 거대분자 뷰어, SAR 부분구조/유사도 및 약물성(Lipinski) 수식, 벡터 SVG 복사, 일괄 스타일 변환, SDF 입출력 지원."
  actions:
    - theme: brand
      text: 시작 가이드
      link: /ko/guide/getting-started
    - theme: alt
      text: Excel 가이드
      link: /ko/guide/excel
    - theme: alt
      text: 수식 레퍼런스
      link: /ko/reference/formula
    - theme: alt
      text: GitHub 저장소
      link: https://github.com/naramdash/func_SMILES2IMG

features:
  - icon: ⚡
    title: 실시간 인셀 수식 & 부분구조 하이라이트
    details: =SMILES2IMG, =INCHI2IMG, =KET2IMG, =NAME2IMG("Aspirin") 수식으로 셀 안에 구조식을 네이티브 삽입하고, "hl:c1ccccc1" 옵션으로 원하는 코어 골격을 하이라이트합니다.
  - icon: 🔬
    title: 화학정보학(SAR) & 약물성(Lipinski) 수식
    details: =SMILES.MATCH, =SMILES.SIMILARITY, =SMILES2LIPINSKI, =SMILES2LOGP, =SMILES2TPSA, =SMILES2EXACTMASS, =SMILES2MW 등 25종 이상의 수식이 0ms 순수 로컬로 작동합니다.
  - icon: ✏️
    title: 초고속 Ketcher 2D 에디터 (<0.05초 재실행)
    details: 엑셀 셀이나 워드/파워포인트 그림을 더블 클릭하면 WebView2 웜 인스턴스 풀링이 적용된 Ketcher v3.18.0 에디터가 즉시 열려 제자리 수정을 지원합니다.
  - icon: 🧬
    title: Mol* 3D 거대분자 뷰어 (Excel · Word · PPT)
    details: 단백질, 핵산, 리간드 복합체를 엑셀, 워드, 파워포인트 모두에서 3D로 시각화합니다. PDB ID 직접 다운로드 및 회전 가능한 GLB 3D 모델·2D 이미지 삽입을 지원합니다.
  - icon: 🖼️
    title: 벡터 SVG 복사 · 다중 서식 요약 · 일괄 스타일 변환
    details: 무손실 벡터 그래픽(Copy SVG) 복사, 분자 전체 물성의 Excel 가로 행(TSV)·마크다운 표·멀티라인 복사, 문서 전체 ACS 흑백·다크 모드 원클릭 일괄 변환(Batch Style)을 지원합니다.
  - icon: 📂
    title: 엑셀 네이티브 SDF 가져오기 / 내보내기
    details: 다중 분자가 담긴 .sdf 파일을 읽어 구조식과 물성 수식이 포함된 엑셀 표로 자동 생성하거나, 시트의 화합물 목록을 .sdf 파일로 내보냅니다.
  - icon: 🛡️
    title: 100% 로컬 오프라인 보안 & 무손실 보존
    details: EPAM Indigo, Ketcher, Mol* 엔진이 로컬에 내장되어 외부 유출 없이 안전하며, 반응식(KET)과 3D 세션이 문서 내부(CustomXMLParts)에 무손실 보존됩니다.
  - icon: 📦
    title: 원클릭 통합 설치 & 인앱 업데이트 확인
    details: 단 한 줄의 PowerShell 명령어로 Excel, Word, PowerPoint 전체에 자동 설치되며, 리본의 [About SMILES] 창에서 최신 버전을 즉시 확인할 수 있습니다.
---

## ⚡ 빠른 설치 (Quick Install)

PowerShell을 열고 아래 명령어를 실행하면 Excel, Word, PowerPoint 전체에 원클릭으로 자동 설치됩니다:

```powershell
# 안정판 릴리즈 (v1.7.0)
irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex
```

::: tip 💡 간편한 업데이트 및 버전 확인
언제든 동일한 설치 명령어를 실행하면 최신 버전으로 자동 업그레이드되며, 오피스 상단 리본의 **[SMILES] ➔ [About SMILES] ➔ [Check for Updates]** 버튼으로도 새 버전을 확인할 수 있습니다!
:::

---

## 🖼️ 오피스 프로그램별 주요 화면

<p align="center">
  <img src="/excel-showcase.png" alt="Microsoft Excel Showcase" width="100%" style="border-radius: 8px; margin-bottom: 16px;" />
</p>

| [📗 Microsoft Excel 가이드](/ko/guide/excel) | [📘 Microsoft Word 가이드](/ko/guide/word) | [📙 Microsoft PowerPoint 가이드](/ko/guide/powerpoint) |
| :---: | :---: | :---: |
| 네이티브 인셀 수식(`=SMILES2IMG`, `=NAME2IMG`), SAR/Lipinski 약물성 스크리닝, `.sdf` 입출력, 일괄 스타일 변환 | 문장 흐름을 따르는 인라인 그림, 표 서식 완벽 보존, Mol* 3D 모델 삽입, 벡터 SVG 복사 및 일괄 스타일 변환 | 슬라이드 정중앙 원클릭 배치, 다크 테마 투명 배경, 회전 가능한 3D GLB 모델, 슬라이드 전체 일괄 스타일 변환 |

| [✏️ Ketcher 2D 에디터](/ko/guide/ketcher) | [🧬 Mol* 3D 뷰어](/ko/guide/molstar) | [📐 수식 레퍼런스](/ko/reference/formula) |
| :---: | :---: | :---: |
| `<0.05초` 초고속 재실행 2D 분자·반응식 스케치, 무손실 KET 캔버스 | Excel · Word · PowerPoint 통합 3D 고분자 뷰어, PDB 다운로드 및 GLB 3D | 25종 이상의 렌더링·물성·약물성·SAR 엑셀 함수 및 옵션 상세 규격 |
