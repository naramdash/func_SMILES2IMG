# 리본 · 우클릭 메뉴 · 진단 도구

Microsoft Excel 상단 **`[SMILES]`** 리본 탭의 전체 명령, 3가지 대상별 **`Chemical Structure`** 우클릭 컨텍스트 메뉴, 그리고 `#VALUE!` 원인을 찾아내는 **`Diagnostics`** 파서 로그 창 레퍼런스입니다.

<FeatureMedia id="excel-ribbon-tools" />

## 1. 상단 `[SMILES]` 리본 전체 명령표

| 리본 그룹 | 버튼 / 드롭다운 메뉴 | Excel에서의 동작 |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit (SMILES)` | 선택한 셀에 대해 Ketcher 2D를 열고 `=SMILES2IMG("...")` 수식을 삽입 (반응식이면 `=KET2IMG` 자동 적용) |
| | `Insert / Edit (InChI)` | 선택한 셀에 대해 Ketcher 2D를 열고 `=INCHI2IMG("...")` 수식을 삽입 |
| | `Insert / Edit (KET)` | 선택한 셀에 대해 Ketcher 2D를 열고 무손실 `=KET2IMG("...")` JSON 수식을 삽입 |
| | `3D Structure (Mol*)` | Mol\* 3D 뷰어 실행 (활성 셀에 4자리 PDB 코드가 적혀 있으면 자동 로드하여 `.glb` 또는 `.png` 삽입) |
| | `Paste Structure` | 클립보드 또는 선택 셀의 SMILES/InChI/KET 텍스트를 0.1초 만에 구조식 수식으로 즉시 변환 |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` | 선택한 수식 셀이나 그림에서 표준 SMILES, IUPAC InChI, KET JSON 추출 복사 |
| | `Copy SVG` | Illustrator, Inkscape, Office 포스터용 무손실 벡터 SVG 그래픽 + XML 코드 복사 |
| | `Molecule Info` | 개별 물성, 17열 TSV·요약·목록·Markdown 복사 또는 17종 상세 물성 창 열기 |
| | `Batch Style` | 선택 영역(단일 셀이면 시트 전체)을 `Standard Color`, `Publication B&W`, `White Background`, `Dark Mode`로 원클릭 통일 |
| | `SDF File` | `.sdf` 파일을 12열 라이브 시트로 가져오기(`Import`) 및 현재 시트를 V2000 `.sdf`로 내보내기(`Export`) |
| **Diagnostics** | `Diagnostics` | `#VALUE!` 수식 오류의 정확한 화학 문법 원인을 기록한 EPAM Indigo 파서 진단 로그 창 열기 |
| **About** | `Online Docs` | 웹 브라우저에서 Excel 공식 문서 카테고리 인트로(`/excel/`) 즉시 열기 |
| | `About func_SMILES2IMG` | 설치 버전, 내장 엔진 버전, 서드파티 라이선스 확인 및 원클릭 최신 업데이트 체크 |

## 2. 우클릭 컨텍스트 메뉴 (3가지 대상 지원)

**일반 셀(`Cell`)**, **엑셀 표(`ListRange`)**, **부동 그림 도형(`Picture`)** 어디서 우클릭하든 3단계 **`Chemical Structure`** 컨텍스트 메뉴로 그리기·복사·물성 조회를 바로 실행할 수 있습니다:

![Excel 3단계 Chemical Structure 우클릭 메뉴](/excel-context-menu-guide.png)

## 3. `Diagnostics` 창으로 `#VALUE!` 파서 오류 진단하기

수식 결과가 `#VALUE!`로 표시될 때 상단 **`[SMILES]` ➔ `[Diagnostics]`** 버튼을 클릭하세요. 닫히지 않은 고리 번호(`unclosed ring`), 괄호 불일치, 잘못된 원자가 등 EPAM Indigo 화학 엔진이 보고한 정확한 문법 오류 메시지를 즉시 확인할 수 있습니다.

## 관련 가이드

- [**일괄 스타일 · TSV · 3D**](/ko/excel/tools-export) — 시트 일괄 스타일 변환, 17열 TSV 복사 및 셀 PDB 기반 3D 단백질 삽입 활용법
- [**SDF 파일 입출력**](/ko/excel/sdf) — 12열 라이브 시트 가져오기 및 V2000 `.sdf` 파일 내보내기
- [**수식 오류**](/ko/excel/troubleshooting/formula-errors) — `#NAME?`, `#VALUE!`, `#N/A` 오류 코드별 원인 및 해결 방법
