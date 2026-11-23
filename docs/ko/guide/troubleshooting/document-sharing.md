# 문서 공유

`func_SMILES2IMG`로 작성한 `.docx`, `.pptx`, `.xlsx` 파일을 **macOS**, 아이패드, 웹 오피스, 또는 추가 기능이 설치되지 않은 PC로 공유할 때의 동작과 권장 방법입니다.

## 1. Microsoft Word (`.docx`) & PowerPoint (`.pptx`)

* **어디서나 100% 정상 표시:** 2D 구조식은 표준 고해상도 PNG 그림으로, 3D 단백질은 표준 `.glb` 3D 모델로 문서에 직접 저장됩니다.
* **열람자는 설치 불필요:** Windows, **macOS**, 아이패드, 웹 오피스에서 추가 기능 없이도 그림 깨짐이나 레이아웃 틀어짐 없이 열람·발표·인쇄·PDF 저장이 가능합니다.
* **왕복 재편집 유지:** 화학 메타데이터(`KET:` / `SMILES:` / `MOLSTAR:`)가 표준 Alt Text와 `CustomXMLParts`에 함께 저장되므로, 나중에 `func_SMILES2IMG`가 설치된 PC에서 다시 열면 더블클릭 편집도 그대로 작동합니다.

## 2. Microsoft Excel (`.xlsx`)

* **캐시된 결과 그대로 표시:** `.FLOAT` 도형, Microsoft 365 인셀 이미지 캐시, 이미 계산된 물성값(`MW`, `Formula`, `LogP`, `TPSA`)은 재계산이 일어나기 전까지 다른 PC에서도 화면에 그대로 보입니다.
* **외부 공유 시 `#NAME?` 예방:** 추가 기능이 없는 동료가 수식 셀을 더블클릭하거나 전체 재계산(<kbd>F9</kbd>)을 누르면 사용자 지정 함수가 `#NAME?`으로 바뀔 수 있습니다.
* **권장 배포 방법:**
  1. 라이브 수식이 담긴 작업용 원본 `.xlsx`는 별도로 보관합니다.
  2. 추가 기능이 없는 외부 수신자에게 보낼 때는 숫자·텍스트 수식 열을 복사(<kbd>Ctrl</kbd> + <kbd>C</kbd>) ➔ **값으로 붙여넣기**하거나, **PDF로 내보내** 전달하는 것이 가장 안전합니다.

## 관련 트러블슈팅

- [**Word 더블클릭 편집**](/ko/word/troubleshooting/double-click) — 공유된 Word/PowerPoint 파일에서 Alt Text 메타데이터가 편집을 유지하는 원리
- [**자주 묻는 질문 (FAQ)**](/ko/guide/faq) — ACS 흑백 변환, 벡터 SVG 복사, .sdf 연동 및 업데이트 안내
