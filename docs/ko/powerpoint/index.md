# PowerPoint 개요

어두운 슬라이드 템플릿 위의 눈부신 하얀 네모 배경, 돌아가지 않는 정적 PyMOL 캡처, 인쇄하면 계단처럼 깨지는 A0 포스터 사진은 이제 그만 쓰세요. Microsoft PowerPoint 안에서 투명 흰색 선화 다크 모드 구조식, 무손실 반응 메커니즘, 발표 중 360° 회전하는 3D 단백질(`Mol*`), 무한 해상도 포스터 벡터 `SVG`까지 한 번에 완성합니다.

<FeatureMedia id="powerpoint-overview" />

## 발표 · 포스터 제작에서 달라지는 3가지 결정적 순간

### 1. 하얀 네모 배경 0% 다크 모드 & 발표 덱 전체 1초 테마 전환 (`Batch Style`)
학회 발표 직전 템플릿 색상이 바뀌었거나, 어두운 학회 덱을 밝은 랩미팅·흑백 유인물로 바꿔야 할 때 도형을 다시 그릴 필요가 없습니다. **프레젠테이션 전체 슬라이드의 2D 구조식**을 클릭 한 번으로 다크 모드, 컬러, ACS 흑백으로 통일합니다.

<FeatureMedia id="powerpoint-slide-batch-style" />

### 2. 발표 중 마우스로 360° 돌려보고 모핑으로 줌인하는 3D 단백질 (`Mol*`)
4자리 PDB 코드(`1CRN`, `6LU7`)나 로컬 `.pdb`/`.cif` 파일을 슬라이드에 **네이티브 3D 모델(`.glb`)** 로 바로 넣습니다. 발표 중 마우스로 결합 포켓을 360° 회전 시연하거나, 동영상 녹화 없이 PowerPoint **모핑(Morph)** 만으로 부드러운 3D 줌인 애니메이션을 연출할 수 있습니다.

<FeatureMedia id="powerpoint-slide-molstar-3d" />

### 3. 4-카드 후보 물질 SAR 그리드 & 텍스트 상자 `0.1초` 즉시 변환
SMILES나 화합물명(`Aspirin`, `Paracetamol`)이 적힌 텍스트 상자를 선택하고 **`[Paste Structure]`** 를 누르면 그 자리에서 비율 고정 구조식으로 바뀌며, 아래에 `Molecule Info` 물성 불릿까지 즉시 붙일 수 있습니다.

<FeatureMedia id="powerpoint-slide-candidate-grid" />

## 기존 방식 vs. `func_SMILES2IMG`

| 발표 · 포스터 작업 | 기존 방식 (ChemDraw / PyMOL / 캡처) | `func_SMILES2IMG` 사용 시 |
| :--- | :--- | :--- |
| **어두운 네이비·슬레이트 슬라이드 배치** | 하얀 네모 배경이 남거나 번거롭게 배경 투명화 작업 필요 | **완전 투명 배경 + 고대비 흰색 결합선(`Batch Style: Dark Mode`)** |
| **발표 전날 템플릿 배경색 변경** | 슬라이드마다 수십 개 이미지를 다시 캡처·정렬 | **빈 배경 클릭 후 `[Batch Style]` 1클릭으로 덱 전체 1초 전환** |
| **단백질 결합 포켓 3D 시연** | 각도별 정적 PyMOL 캡처 여러 장 또는 무거운 MP4 녹화 | **슬라이드 쇼에서 360° 회전하는 `3D Model (GLB)` + 모핑 줌인** |
| **학회장 연단 PC · macOS 발표** | ChemDraw 개체 깨짐 또는 폰트 누락 위험 | **표준 PNG + `.glb` 내장 — 애드인 없는 PC/Mac에서도 100% 동작** |
| **1미터 대형 A0 학회 포스터 인쇄** | 비트맵 이미지 확대 시 결합선·글자가 계단처럼 흐려짐 | **`[Copy SVG]`로 어떤 배율에서도 칼날처럼 선명한 벡터 패스 복사** |

## 발표장과 연구실에서 안심하고 쓸 수 있는 이유

<OfficeOverviewCta app="powerpoint" />

## 설치 후 1분 안에 체험하기

1. **텍스트 상자 `0.1초` 변환:** 슬라이드에 텍스트 상자를 넣고 `Aspirin` 또는 `CC(=O)Oc1ccccc1C(=O)O`를 적은 뒤 상자 테두리를 선택해 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 클릭해 보세요.
2. **다크 모드 & 더블클릭 수정:** 어두운 슬라이드라면 **`[Batch Style]` ➔ `[Dark Mode]`** 를 누르고, 구조식을 **더블클릭**해 **Ketcher 2D**에서 치환기를 수정한 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>로 제자리 업데이트를 확인하세요.
3. **8슬라이드 공식 Playbook(`.pptx`) 실습:** <a href="/func_SMILES2IMG/examples/func_SMILES2IMG_PowerPoint_Playbook.pptx" download="func_SMILES2IMG_PowerPoint_Playbook.pptx"><strong>func_SMILES2IMG_PowerPoint_Playbook.pptx</strong></a>를 내려받아 하단 **발표자 노트**를 켠 채 0.1초 텍스트 상자 변환, 4카드 그리드, KET 반응식, Mol\* `1CRN` 슬라이드를 순서대로 따라 해 보세요 ([8슬라이드 상세 구성 →](/ko/powerpoint/playbook)).

## 작업별 실전 가이드 바로가기

| PowerPoint에서 하려는 작업 | 가이드 | 핵심 명령 |
| :--- | :--- | :--- |
| 학회 3D 모션 발표 · 랩미팅 · A0 포스터 · 단계별 반응식 애니메이션 보기 | [PowerPoint 활용 사례](/ko/powerpoint/use-cases) | 5가지 발표·포스터 실전 시나리오 |
| 텍스트 상자 제자리 `0.1초` 변환, 2D 스케치, 무손실 반응 메커니즘(`KET`) 제작 | [1. 2D 구조식 · 반응식 슬라이드](/ko/powerpoint/slides-structures) | `[Paste Structure]` / `[Insert / Edit Structure]` |
| 하얀 네모 없는 다크 모드 선화, 4-카드 SAR 그리드, 덱 전체 테마 1초 전환 | [2. 다크 모드 · 후보 카드 · 덱 테마 전환](/ko/powerpoint/dark-mode) | `[Batch Style]` + `[Molecule Info]` |
| 슬라이드 쇼 중 360° 회전하는 3D 단백질(`.glb`) 및 **모핑(Morph)** 줌인 연출 | [3. 360° 3D 단백질 발표 (`Mol*` & 모핑)](/ko/powerpoint/3d-proteins) | `[3D Structure (Mol*)]` |
| A0 포스터 벡터 `SVG`, 발표 전 Q&A 노트 준비, 학회장 PC/Mac 점검 | [4. A0 포스터(SVG) · Q&A · 발표장 공유](/ko/powerpoint/publishing-qa) | `[Copy SVG]` / `[Molecule Info]` |
| 상단 `[SMILES]` 리본 전체 기능 및 5종 슬라이드 우클릭 메뉴 살펴보기 | [리본 · 우클릭 메뉴](/ko/powerpoint/reference) | 상단 `[SMILES]` 탭 & 우클릭 |
