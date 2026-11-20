# 구조식 삽입 · 텍스트 변환

오프라인 **Ketcher 2D 편집기**로 Word 본문 커서 위치에 구조식을 인라인으로 삽입하거나, 본문·클립보드의 SMILES 텍스트를 0.1초 만에 구조식 그림으로 바꿉니다.

## 1. 본문에 인라인 구조식 그리기 (`Ctrl + Shift + K`)

<FeatureMedia id="word-ch1-insert-ketcher" />

1. 구조식을 넣을 본문 위치에 커서를 둡니다.
2. 상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (또는 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>)를 누릅니다.
3. Ketcher에서 분자를 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 본문 줄 안에 깔끔하게 삽입됩니다.

## 2. 선택 텍스트 · 클립보드 SMILES 즉시 변환 (`Paste Structure`)

<FeatureMedia id="word-ch2-quick-insert" />

1. Word 본문 안의 SMILES 문자열(예: `Cn1cnc2c1c(=O)n(C)c(=O)n2C`)을 드래그해 선택합니다.
2. 상단 **`[SMILES]` ➔ `[Paste Structure]`** (또는 우클릭 ➔ **`Chemical Structure` ➔ `Paste Structure`**)를 클릭합니다.
3. 선택한 텍스트가 가로세로 비율이 고정된 인라인 구조식으로 즉시 바뀝니다. *(텍스트를 선택하지 않은 상태면 클립보드에 복사된 SMILES/InChI/KET를 변환해 삽입합니다.)*
