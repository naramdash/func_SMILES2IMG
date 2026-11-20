# 더블클릭 편집

Microsoft Word에서 화학 구조식을 더블클릭했을 때 **Ketcher 2D**나 **Mol\* 3D** 편집기가 열리지 않는 원인과 해결 방법입니다.

## 1. Word가 편집 가능한 구조식을 인식하는 방식

`func_SMILES2IMG`로 삽입한 구조식은 그림의 **대체 텍스트**(Alt Text)에 화학 메타데이터(`SMILES:...`, `KET:...`, `MOLSTAR:...`)를 기록하고, 정밀 좌표와 3D 세션은 Word 문서의 `CustomXMLParts`에 무손실 저장합니다.

**텍스트 줄 안 그림**(`InlineShape`)과 **부동 그림 · 3D 모델**(`Shape`) 모두 더블클릭 제자리 편집을 지원합니다.

## 2. 더블클릭이 동작하지 않는 주요 원인

| 원인 | 발생 이유 | 해결 방법 |
| :--- | :--- | :--- |
| **일반 스크린샷 · 외부 붙여넣기** | 캡처 도구나 PDF, 외부 앱에서 복사한 일반 그림에는 화학 메타데이터가 없습니다. | **`[SMILES]` ➔ `[Insert / Edit Structure]`** 또는 **`[Paste Structure]`** 버튼으로 삽입하세요. |
| **대체 텍스트(Alt Text) 삭제** | 다른 매크로나 수동 편집으로 `SMILES:`, `KET:`, `MOLSTAR:` 접두어가 지워졌습니다. | 그림 우클릭 ➔ **대체 텍스트 보기**에서 태그가 남아 있는지 확인하세요. |
| **Word COM 추가 기능 비활성화** | 더블클릭 가로채기는 Word COM 추가 기능이 로드되어 있어야 동작합니다. | [리본 탭 미표시](/ko/guide/troubleshooting/ribbon-tab)(`COM 추가 기능` / `사용할 수 없는 항목`)를 확인하세요. |

## 3. 단축키 · 리본으로 바로 열기

Word에서 구조식 그림을 선택한 상태에서 아래 방법으로 즉시 편집기를 열 수 있습니다:

* <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> — Ketcher 2D 즉시 열기
* 상단 리본 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (2D) 또는 **`[3D Structure (Mol*)]`** (3D)
* 그림 **우클릭 ➔ `Chemical Structure` ➔ `Edit Structure...`**

## 관련 가이드

- [**표 레이아웃 · 스타일**](/ko/word/troubleshooting/table-layout) — 재편집 시 표 열 너비 유지 및 논문 전체 ACS 흑백 일괄 변환 팁
- [**문서 공유**](/ko/guide/troubleshooting/document-sharing) — Mac이나 추가 기능이 없는 PC에서 Word .docx 구조식이 보이는 방식
