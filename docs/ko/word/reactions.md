# 합성 반응식 (KET)

반응 화살표(`➔`), `+` 기호, 촉매·온도·수율 주석이 포함된 다단계 합성 반응식을 편집 가능한 무손실 형태로 Word 문서에 삽입합니다.

<FeatureMedia id="word-ch4-reaction-scheme" />

## Word에서 무손실 KET 반응식이 저장되는 방식

1. **`[SMILES]` ➔ `[Insert / Edit Structure]`** (<kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>)로 Ketcher를 엽니다.
2. 반응물, 화살표(`➔`), `+`, 촉매 및 반응 조건 텍스트를 그립니다.
3. **`Apply to Document`** (<kbd>Ctrl</kbd> + <kbd>Enter</kbd>)를 눌러 문서에 삽입합니다.
4. 전체 KET JSON 좌표가 그림의 `AlternativeText`와 Word `CustomXMLParts`에 함께 저장되므로, `.docx`를 저장·공유한 뒤에도 `func_SMILES2IMG`가 설치된 PC에서 **더블클릭**하면 원본 화살표와 배치 그대로 다시 열립니다.
