# SMILES2IMG — Excel 분자 구조식 자동 렌더러

> **SMILES 화학 구조식 문자열을 Excel 셀 안의 고해상도 분자 이미지로 즉시 변환하는 Excel 추가 기능입니다.**  
> 외부 웹 서버나 클라우드 전송 없이, PC 로컬에서 100% 오프라인으로 안전하고 빠르게 렌더링됩니다.

---

## 🚀 빠른 시작 (설치 방법)

Windows 데스크톱 Excel(Microsoft 365, Office 2021/2024 등 셀 이미지 기능 지원 버전)에서 사용하실 수 있습니다.

### 1단계: 내 Excel 비트수(32비트 vs 64비트) 확인
1. Excel 실행 후 **[파일] → [계정] → [Excel 정보]** 클릭
2. 팝업 창 첫 줄 끝부분에서 **`32비트`** 또는 **`64비트`** 확인

### 2단계: 추가 기능(XLL) 파일 준비
내 Excel 비트수에 맞는 **`.xll` 파일 1개만** 원하는 위치(예: 바탕화면, `C:\ExcelAddIns\` 등)에 복사합니다:
* **64비트 Excel:** [`cs_SMILES2IMG/dist/x64/Smiles2Img-AddIn64-packed.xll`](cs_SMILES2IMG/dist/x64/Smiles2Img-AddIn64-packed.xll)
* **32비트 Excel:** [`cs_SMILES2IMG/dist/x86/Smiles2Img-AddIn-packed.xll`](cs_SMILES2IMG/dist/x86/Smiles2Img-AddIn-packed.xll)

> [!NOTE]
> **단일 파일(Single-File) 100% 자체 내장:**  
> 분자 렌더링 C++ 네이티브 엔진 및 수식 자동완성 라이브러리가 `.xll` 파일 하나 안에 모두 압축 내장되어 있습니다.  
> 다른 DLL 파일들을 함께 옮길 필요 없이, **오직 `.xll` 파일 단 1개만** 있으면 즉시 작동합니다.

### 3단계: Excel에 추가 기능 등록
1. Excel 메뉴에서 **[파일] → [옵션] → [추가 기능]** 이동
2. 하단 **[관리: Excel 추가 기능]** 옆의 **[이동(G)...]** 버튼 클릭
3. **[찾아보기(B)...]**를 눌러 비트수에 맞는 XLL 파일 선택:
   * 64비트: `Smiles2Img-AddIn64-packed.xll`
   * 32비트: `Smiles2Img-AddIn-packed.xll`
4. 목록에 `Smiles2Img-AddIn`이 체크된 것을 확인하고 **[확인]** 클릭

---

## 💡 사용 방법

Excel 워크시트의 셀에 `=SMILES2IMG(...)` 수식을 입력하면 즉시 셀 안에 분자 이미지가 삽입됩니다.

```excel
=SMILES2IMG(A2)
```

> [!TIP]
> • **실시간 수식 도움말(IntelliSense):** 셀에 `=SMILES2IMG(` 를 입력하면 엑셀 기본 함수처럼 `SMILES2IMG(smiles, [color], [options])` 인수 안내 풍선도움말과 설명이 자동으로 표시됩니다.  
> • 분자 구조가 잘 보이도록 행 높이(예: 80~120pt)와 열 너비(예: 25~40)를 넉넉하게 늘려주시면 훨씬 보기 좋습니다.

---

## 📖 함수 구문 및 인자 안내

```excel
=SMILES2IMG(smiles, [background], [color], [transform])
```

| 순서 | 인자명 | 필수 여부 | 기본값 | 설명 및 허용 값 |
| :---: | :--- | :---: | :---: | :--- |
| **1** | **`smiles`** | **필수** | - | SMILES 문자열 또는 해당 문자열이 들어있는 셀 참조 (예: `A2`, `"CCO"`) |
| **2** | **`background`** | 선택 | `"white"` | **배경 색상 설정:**<br>• `"trans"` 또는 `"transparent"`, `"nobg"`: 투명 배경 (셀 배경 투과)<br>• `"white"`: 흰색 불투명 배경 (기본값)<br>• `"#RRGGBB"`: 사용자 지정 16진수 색상 (예: `"#FFFF00"`, `"#FFF8DC"`) |
| **3** | **`color`** | 선택 | `TRUE` | **원소 색상 모드 (학술 논문/인쇄용):**<br>• `TRUE`: 산소(빨강), 질소(파랑), 황(황갈색) 등 원소별 표준 컬러 (기본값)<br>• `FALSE`: 흑백(Grayscale) 모드 |
| **4** | **`transform`** | 선택 | `""` | **기하학적 변환 (회전 및 반전):**<br>• `flip`: 좌우 반전 (교과서/논문 표준 배치로 전환)<br>• `flipy`: 상하 반전<br>• `rot90`, `rot180`, `rot270`: 시계 방향 회전 |

---

## 🎨 주요 활용 예시

### 1. 기본 컬러 렌더링
```excel
=SMILES2IMG(A2)
```

### 2. 셀 배경색이 투과되는 투명 배경
엑셀 셀에 채우기 색상(노란색, 연회색 등)이나 표 서식이 적용되어 있을 때 유용합니다. **(2번째 인자에 바로 입력)**
```excel
=SMILES2IMG(A2, "trans")
=SMILES2IMG(A2, "transparent")
```

### 3. 사용자 지정 배경색
```excel
=SMILES2IMG(A2, "#FFFF00")              ' 노란색 배경
=SMILES2IMG(A2, "#FFF8DC")              ' 크림색 배경
```

### 4. 학술 논문/인쇄용 ACS 스타일 흑백 표기
```excel
=SMILES2IMG(A2, , FALSE)                ' 기본 흰 배경 + 흑백
=SMILES2IMG(A2, "trans", FALSE)         ' 투명 배경 + 흑백
```

### 5. 교과서/논문 표준 도식에 맞춘 좌우 반전 (`flip`)
SMILES 기입 순서로 인해 고리 위치가 교과서 도식과 반대로 렌더링될 때(예: 니코틴, 티아민) 즉시 표준 방향으로 뒤집을 수 있습니다.
```excel
=SMILES2IMG(A2, , , "flip")             ' 기본 배경 + 좌우 반전
=SMILES2IMG(A2, "trans", , "flip")      ' 투명 배경 + 좌우 반전
=SMILES2IMG(A2, "trans", FALSE, "flip") ' 투명 배경 + 흑백 + 좌우 반전
```

### 6. 회전 조절
```excel
=SMILES2IMG(A2, , , "rot90")            ' 시계 방향 90도 회전
=SMILES2IMG(A2, "trans", , "rot90")     ' 투명 배경 + 90도 회전
```

---

## 📂 프로젝트 구조 및 개발 문서

| 폴더 | 구현 기술 | 주요 특징 | 문서 링크 |
| :--- | :--- | :--- | :--- |
| **`cs_SMILES2IMG/`** | **C# (.NET 4.8) + Excel-DNA + EPAM Indigo** | **[메인 추천 버전]** 네이티브 데스크톱 속도, ACS 표준 렌더링 최적화, 고해상도 가변 크기, 100% 로컬 오프라인 실행 | [상세 문서 보기](cs_SMILES2IMG/README.md) |
| **`js_SMILES2IMG/`** | TypeScript + Office.js + SmilesDrawer | Office.js 웹 런타임 기반 추가 기능 (웹/클라우드 환경 대응 연구용) | [상세 문서 보기](js_SMILES2IMG/README.md) |

---

## 🛠️ 문제 해결 (FAQ)

* **Q. 수식 결과가 `#VALUE!`로 나옵니다.**
  * 참조한 셀에 올바른 SMILES 문자열이 들어있는지 확인하세요.
  * 수식이 참조하는 셀에 **셀 병합(Merged Cells)**이 되어 있으면 빈 셀을 참조하여 `#VALUE!`가 발생할 수 있습니다.
* **Q. 셀 이미지가 너무 작거나 잘려 보입니다.**
  * 이미지는 셀 크기에 맞춰 자동 축척됩니다. 행 높이와 열 너비를 충분히 늘려주세요.
* **Q. 상세 로그나 오류를 확인하고 싶습니다.**
  * Excel 상단 리본 메뉴의 **[SMILES] → [Show diagnostics]** 창에서 상세 로그를 확인할 수 있습니다.
