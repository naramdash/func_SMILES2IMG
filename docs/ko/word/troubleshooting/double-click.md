# 더블클릭 · 재편집 문제 해결

Microsoft Word에서 구조식을 더블클릭했을 때 **Ketcher 2D**나 **Mol\* 3D** 대신 Word 기본 **그림 서식** 창이 열릴 때의 원인과 해결 방법, 그리고 공동저자 문서 공유 시 메타데이터 유지 방법을 정리했습니다.

## 1. 더블클릭 시 Ketcher · Mol\* 대신 Word 그림 서식 창이 열림 {#double-click}

`func_SMILES2IMG` 로 삽입한 그림은 **대체 텍스트(Alt Text)** 에 `SMILES:`, `KET:`, `MOLSTAR:` 태그를 저장하며, Mol\* 3D 세션은 문서 내부 `CustomXMLParts` 에 함께 보관합니다. 더블클릭이 가로채지지 않을 때는 아래 항목을 순서대로 확인하세요:

* **일반 스크린샷 · PDF 캡처 그림:** 외부 앱에서 복사해 온 비트맵 이미지에는 화학 메타데이터가 없습니다. **`[SMILES]` ➔ `[Insert / Edit Structure]`** 나 **`[Paste Structure]`** 로 다시 삽입하세요.
* **Alt Text 삭제:** 다른 매크로나 수동 편집 과정에서 대체 텍스트가 지워졌을 수 있습니다. 그림 **우클릭 ➔ 대체 텍스트 보기**에서 `SMILES:`, `KET:`, `MOLSTAR:` 접두어가 남아 있는지 확인하세요.
* **Word COM 추가 기능 비활성화:** 상단 리본에 `[SMILES]` 탭 자체가 보이지 않는다면 **파일 ➔ 옵션 ➔ 추가 기능 ➔ COM 추가 기능 / 사용할 수 없는 항목**에서 격리를 해제하세요 ([리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)).
* **부동 그림 전환 직후 이벤트 누락:** 텍스트 배치(*줄 안* ↔ *사각형*)를 바꾼 직후 일시적으로 반응이 없다면 리본/우클릭 메뉴로 한 번 열어 재저장하세요. *(참고: 더블클릭 훅은 Word 본문 문서 캔버스(`_WwG`) 클라이언트 영역에서만 동작하므로 리본·작업 창 더블클릭에는 반응하지 않습니다.)*

## 2. Mac 공동저자 공유 · 메타데이터 유지 {#sharing}

* **추가 기능 없는 Mac · PC 열람:** 삽입된 2D PNG와 3D `.glb` 모델은 Word 표준 개체이므로, 열람자는 추가 기능 설치 없이 Windows·**macOS**·웹 Word에서 100% 동일하게 볼 수 있습니다.
* **왕복 재편집 유지:** 화학 메타데이터가 표준 **Alt Text** 와 `CustomXMLParts` 에 저장되므로, Mac을 쓰는 공동저자가 본문을 수정해 돌려보내더라도 Windows에서 다시 열면 더블클릭 편집이 그대로 동작합니다 ([문서 공유 가이드](/ko/guide/troubleshooting/document-sharing)).
* **대형 인쇄 · 일러스트레이터 편집:** 그림 선택 후 **`[SMILES]` ➔ `[Copy SVG]`** 로 벡터 패스를 복사해 Adobe Illustrator나 Inkscape에 바로 붙여넣을 수 있습니다.

## 3. 더블클릭 외 대체 열기 방법

다른 추가 기능 단축키와 겹치더라도 아래 두 경로로 언제든 제자리 재편집을 실행할 수 있습니다:

1. 그림을 선택한 상태에서 상단 리본 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) 또는 **`[3D Structure (Mol*)]`** (3D) 클릭
2. 그림 **우클릭 ➔ `Chemical Structure` ➔ `Insert / Edit Structure...`** (2D) 또는 **`3D Structure (Mol*)...`** (3D) 클릭

## 관련 가이드

- [**표 잘림 · 일괄 스타일**](/ko/word/troubleshooting/table-layout) — 표 셀 상단 그림 잘림, 열 너비 밀림, `Batch Style` 적용 범위 해결
- [**리본 · 우클릭 메뉴**](/ko/word/reference) — Word `[SMILES]` 리본 전체 명령표 및 5종 우클릭 메뉴 지도
- [**리본 탭 미표시**](/ko/guide/troubleshooting/ribbon-tab) · [**문서 공유**](/ko/guide/troubleshooting/document-sharing) — 비활성화된 Word 추가 기능 복구 및 공동저자 호환성 안내
