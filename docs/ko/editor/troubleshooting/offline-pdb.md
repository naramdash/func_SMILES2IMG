# 오프라인 PDB 파일

사내 방화벽이 RCSB PDB / PDBe 외부 접속을 차단할 때 **Mol\* 3D**에서 3D 단백질 구조를 오프라인으로 여는 방법입니다.

## 1. 4자리 PDB 코드 조회가 실패하는 이유

Mol\* 상단에 4자리 PDB 코드(`1CRN`, `4HHB`, `6LU7` 등)를 입력하고 **`Fetch`** 버튼을 누르면 HTTPS로 **RCSB PDB / PDBe**에서 공개 결정 구조를 내려받습니다.

연구실·사내 방화벽이 외부 HTTPS 요청을 제한하는 경우 원격 다운로드가 시간 초과되거나 실패할 수 있습니다.

## 2. 로컬 3D 구조 파일을 오프라인으로 열기

Mol\* 3D 뷰어 런타임 자체는 `func_SMILES2IMG` 내부에 내장되어 있어 로컬 파일은 **인터넷 연결 없이 열립니다**:

1. 사용할 구조 파일(`.pdb`, `.cif`, `.bcif`, `.ent`, `.mol`, `.sdf`, `.molj`, `.molx`)을 PC에 준비합니다.
2. 리본에서 **`[SMILES]` ➔ `[3D Structure (Mol*)]`** 메뉴를 클릭합니다.
3. Mol\* 상단 툴바의 **`Open File...`** 버튼을 눌러 로컬 파일을 선택합니다.
4. 원하는 스타일(Cartoon, Surface 등)을 설정한 뒤 **`Apply to Slide`** / **`Apply to Document`** (또는 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>)를 누릅니다.

<FeatureMedia id="common-molstar-3d" />

## 3. 오프라인 모델 추후 재편집

카메라 각도, 스타일 레시피(`.molj`), 모델 정보는 문서 메타데이터(`CustomXMLParts`)에 함께 저장되므로 나중에 더블클릭하면 같은 화면 그대로 다시 열립니다.

## 관련 가이드

- [**WebView2 로딩 오류**](/ko/editor/troubleshooting/webview2-loading) — Ketcher/Mol* 빈 화면 해결 및 %LOCALAPPDATA%\Smiles2Img 캐시 초기화
- [**Mol\* 3D 단백질 뷰어 가이드**](/ko/editor/molstar) — 3D GLB 모델 삽입 및 논문용 2D 단백질 이미지 렌더링 3단계 안내
