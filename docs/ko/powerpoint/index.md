# PowerPoint 개요

Microsoft PowerPoint 안에서 투명 배경 다크 모드 구조식, 반응 메커니즘, 360° 회전하는 3D 단백질(`Mol*`), 확대해도 깨지지 않는 포스터용 벡터 `SVG` 복사를 지원합니다.

<FeatureMedia id="powerpoint-overview" />

## 주요 기능

### 1. 투명 배경 다크 모드 & 프레젠테이션 일괄 테마 전환 (`Batch Style`)
학회 발표 템플릿 색상이 바뀌었거나, 어두운 덱을 밝은 랩미팅·흑백 유인물로 바꿔야 할 때 도형을 다시 그릴 필요가 없습니다. **프레젠테이션 전체 슬라이드의 2D 구조식**을 `[Batch Style]` 메뉴로 다크 모드, 컬러, ACS 흑백으로 일괄 전환합니다.

<FeatureMedia id="powerpoint-slide-batch-style" />

### 2. 발표 중 360° 회전 & 모핑 줌인 3D 단백질 (`Mol*`)
4자리 PDB 코드(`1CRN`, `6LU7`)나 로컬 `.pdb`/`.cif` 파일을 슬라이드에 **네이티브 3D 모델(`.glb`)** 로 넣습니다. 발표 중 마우스로 결합 포켓을 360° 회전 시연하거나, 별도 영상 녹화 없이 PowerPoint **모핑(Morph)** 으로 3D 줌인 전환을 구성할 수 있습니다.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

### 3. 4-카드 후보 물질 SAR 그리드 & 텍스트 상자 구조 변환
SMILES나 화합물명(`Aspirin`, `Paracetamol`)이 적힌 텍스트 상자를 선택하고 **`[Paste Structure]`** 를 누르면 그 자리에서 비율 고정 구조식으로 변환되며, 아래에 `Molecule Info` 물성 요약을 붙일 수 있습니다.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

## 작업 방식 비교

| 발표 · 포스터 작업 | 기존 방식 (ChemDraw / PyMOL / 캡처) | `func_SMILES2IMG` 사용 시 |
| :--- | :--- | :--- |
| **어두운 네이비·슬레이트 슬라이드 배치** | 흰색 사각형 배경이 남거나 별도 배경 제거 작업 필요 | **투명 알파 채널 배경 + 흰색 결합선(`Batch Style: Dark Mode`)** |
| **템플릿 배경색 변경 시 스타일 통일** | 슬라이드마다 이미지를 다시 캡처·정렬 | **빈 배경 클릭 후 `[Batch Style]`로 덱 전체 일괄 전환** |
| **단백질 결합 포켓 3D 시연** | 각도별 정적 PyMOL 캡처 또는 MP4 녹화 | **슬라이드 쇼에서 360° 회전하는 `3D Model (GLB)` + 모핑 줌인** |
| **학회장 연단 PC · macOS 발표** | 전용 개체 호환 문제 또는 폰트 누락 위험 | **표준 PNG + `.glb` 내장 — 애드인 없는 PC/Mac에서도 동일 표시** |
| **대형 A0 학회 포스터 인쇄** | 비트맵 이미지 확대 시 결합선·글자가 흐려짐 | **`[Copy SVG]`로 확대해도 깨지지 않는 벡터 패스 복사** |

## 발표장 및 로컬 실행 환경

<OfficeOverviewCta app="powerpoint" />

## 설치 후 기본 체험

1. **텍스트 상자 변환:** 슬라이드에 텍스트 상자를 넣고 `Aspirin` 또는 `CC(=O)Oc1ccccc1C(=O)O`를 적은 뒤 상자 테두리를 선택해 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 클릭해 보세요.
2. **다크 모드 & 더블클릭 수정:** 어두운 슬라이드라면 **`[Batch Style]` ➔ `[Dark Mode]`** 를 누르고, 구조식을 **더블클릭**해 **Ketcher 2D**에서 치환기를 수정한 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>로 제자리 업데이트를 확인하세요.
3. **8슬라이드 공식 Playbook(`.pptx`) 실습:** <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_PowerPoint_Playbook.pptx" download="func_SMILES2IMG_PowerPoint_Playbook.pptx"><strong>func_SMILES2IMG_PowerPoint_Playbook.pptx</strong></a>를 내려받아 하단 **발표자 노트**를 켠 채 텍스트 상자 변환, 4카드 그리드, KET 반응식, Mol\* `1CRN` 슬라이드를 순서대로 따라 해 보세요 ([8슬라이드 상세 구성 →](/ko/powerpoint/playbook)).

## 작업별 가이드 바로가기

| PowerPoint에서 하려는 작업 | 가이드 | 핵심 명령 |
| :--- | :--- | :--- |
| 학회 3D 모션 발표 · 랩미팅 · A0 포스터 · 단계별 반응식 애니메이션 보기 | [PowerPoint 활용 사례](/ko/powerpoint/use-cases) | 5가지 발표·포스터 시나리오 |
| 텍스트 상자 제자리 변환, 2D 스케치, 무손실 반응 메커니즘(`KET`) 제작 | [1. 2D 구조식 · 반응식 슬라이드](/ko/powerpoint/slides-structures) | `[Paste Structure]` / `[Insert / Edit Structure]` |
| 투명 배경 다크 모드 선화, 4-카드 SAR 그리드, 덱 전체 테마 전환 | [2. 다크 모드 · 후보 카드 · 덱 테마 전환](/ko/powerpoint/dark-mode) | `[Batch Style]` + `[Molecule Info]` |
| 슬라이드 쇼 중 360° 회전하는 3D 단백질(`.glb`) 및 **모핑(Morph)** 줌인 연출 | [3. 360° 3D 단백질 발표 (`Mol*` & 모핑)](/ko/powerpoint/3d-proteins) | `[3D Structure (Mol*)]` |
| A0 포스터 벡터 `SVG`, 발표 전 Q&A 노트 준비, 학회장 PC/Mac 점검 | [4. A0 포스터(SVG) · Q&A · 발표장 공유](/ko/powerpoint/publishing-qa) | `[Copy SVG]` / `[Molecule Info]` |
| 상단 `[SMILES]` 리본 전체 기능 및 5종 슬라이드 우클릭 메뉴 살펴보기 | [리본 · 우클릭 메뉴](/ko/powerpoint/reference) | 상단 `[SMILES]` 탭 & 우클릭 |
