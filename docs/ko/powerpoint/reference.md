# PowerPoint 리본 · 우클릭 메뉴

Microsoft PowerPoint 상단 **`[SMILES]`** 리본 탭 전체 명령표와 5가지 슬라이드 상황별 **`Chemical Structure`** 우클릭 컨텍스트 메뉴 레퍼런스입니다.

<FeatureMedia id="powerpoint-slide-batch-style" />

## 1. 상단 `[SMILES]` 리본 전체 명령표

| 리본 그룹 | 버튼 / 드롭다운 메뉴 | PowerPoint에서의 동작 |
| :--- | :--- | :--- |
| **Insert Structure** | `Insert / Edit Structure` | Ketcher 2D를 열어 슬라이드에 새 2D 구조식을 삽입하거나(텍스트 상자 선택 시 자동 로드) 선택한 구조식 도형을 재편집 |
| | `3D Structure (Mol*)` | Mol\* 3D를 열어 360° 회전 `3D Model (GLB)` 또는 투명 `2D Picture` 삽입(PDB 코드 선택 시 자동 로드) 및 재편집 |
| | `Paste Structure` | 선택한 텍스트 상자(SMILES·InChI·KET·화합물명/CAS)를 `0.1초` 만에 그 위치에서 구조식으로 치환 (클립보드 변환도 지원) |
| **Structure Tools** | `Copy SMILES` / `Copy InChI` / `Copy KET` | 선택한 슬라이드 구조식(또는 텍스트)에서 표준 SMILES, IUPAC InChI, 무손실 KET JSON(`Shape.Tags["KET"]`) 복사 |
| | `Copy SVG` | 대형 A0 학회 포스터, Adobe Illustrator, Inkscape용 무손실 벡터 SVG 그래픽 + XML 텍스트 복사 |
| | `Molecule Info` | 개별 물성, Q&A/카드용 줄바꿈 목록(`Multi-Line`), 17열 요약표 복사 및 17종 상세 물성 창 열기 |
| | `Batch Style` | 선택한 도형(선택이 없으면 **프레젠테이션 전체 슬라이드**의 비그룹 2D 구조식)을 4종 프리셋으로 1초 일괄 전환 |
| **About** | `Online Docs` | 웹 브라우저에서 PowerPoint 공식 문서 카테고리 인트로(`/powerpoint/`) 즉시 열기 |
| | `About func_SMILES2IMG` | 설치 버전, 내장 엔진 버전, 서드파티 라이선스 확인 및 원클릭 최신 업데이트 체크 |

## 2. 우클릭 컨텍스트 메뉴 (5가지 슬라이드 상황)

슬라이드의 **그림 개체(`ContextMenuPicture`)**, **도형(`ContextMenuShape`)**, **텍스트 상자 편집 상태(`ContextMenuText`)**, **표(`ContextMenuTable`)** 에서 우클릭하면 3단계 **`Chemical Structure`** 전체 메뉴가 나타나며, **빈 슬라이드 여백(`ContextMenuFrame`)** 에서는 빠른 삽입 3종 메뉴(`Insert / Edit`, `3D Structure (Mol*)`, `Paste Structure`)가 표시됩니다:

![PowerPoint 3단계 Chemical Structure 우클릭 메뉴](/powerpoint-context-menu-guide.png)

## 관련 가이드

- [**다크 슬라이드 · 그룹 도형 문제 해결**](/ko/powerpoint/troubleshooting/dark-slides) — 하얀 네모 배경, 어두운 선화, 그룹화(`Ctrl+G`) 누락 해결
- [**3D 모델 · 더블클릭 · 발표장 호환성**](/ko/powerpoint/troubleshooting/3d-editing) — `.glb` 지원·용량 제한, 더블클릭 미작동, 학회장 PC/Mac 점검
- [**1. 2D 구조식 · 반응식 슬라이드**](/ko/powerpoint/slides-structures) · [**2. 다크 모드 · 후보 카드 · 덱 테마 전환**](/ko/powerpoint/dark-mode) — 슬라이드 작성 및 테마 변환 가이드
- [**오프라인 PDB 파일 (Mol\*)**](/ko/editor/troubleshooting/offline-pdb) — 방화벽 환경에서 로컬 `.pdb` / `.cif` 파일 열기
