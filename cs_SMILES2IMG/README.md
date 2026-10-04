# SMILES2IMG — C# (Excel-DNA + Indigo.Net)

> 📌 **일반 Excel 사용자를 위한 빠른 설치 및 사용법 안내는 [루트 README.md](../README.md)를 참고하세요.**  
> 본 문서는 C# 기반 Excel-DNA 추가 기능의 화학 렌더링 세부 구현 원리, 함수 옵션 규격, 빌드 및 개발자 검증에 대해 다룹니다.

---

## 1. 개요 및 아키텍처

- **기반 기술:** [.NET Framework 4.8](https://dotnet.microsoft.com/), [Excel-DNA 1.9.0](https://excel-dna.net/), [ExcelDna.IntelliSense 1.9.0](https://github.com/Excel-DNA/IntelliSense), [EPAM Indigo.Net 1.48.0-rc.1](https://www.nuget.org/packages/Indigo.Net/1.48.0-rc.1)
- **실시간 수식 툴팁:** `ExcelDna.IntelliSense`가 내장되어 `=SMILES2IMG(` 타이핑 시 내장 함수와 동일한 인자 안내 풍선도움말 실시간 제공
- **실행 환경:** Excel 32비트 및 64비트 (네이티브 XLL 배포)
- **보안 및 오프라인:** 외부 HTTP 요청, 웹 브라우저 엔진(WebView2), Node.js 서버가 일절 필요 없는 100% 로컬 인메모리 파이프라인

```text
[Excel 셀 수식] =SMILES2IMG(A2, ...)
       │
       ▼
[Functions.cs] 인수 정규화 & 파싱 (Color, Transform, Background)
       │
       ▼
[CellImages.cs] 통합문서 내 숨김 캐시 시트(__SMILES2IMG) 해시 키 조회 (MATCH)
  ├─ 캐시 히트(Hit)   ──> 기존 이미지 셀의 ExcelReference 즉시 반환 (0ms)
  └─ 캐시 미스(Miss)  ──> MoleculeRenderer로 고화질 PNG 생성
                              │
                              ▼
                     [Indigo.Net C++ 엔진]
                     • 고리 가교/접합부 수소 선별 보존
                     • Kekulé 비방향족화(dearomatize)
                     • 스마트 수평 레이아웃 & 2D 좌표 변환
                     • 헤테로원자 결합 메틸기 텍스트 라벨링
                              │
                              ▼
                     [CellImageUpdates.cs]
                     비동기 매크로 큐를 통해 __SMILES2IMG 시트에 로컬 이미지 삽입 후 수식 셀 갱신
```

---

## 2. 화학 구조 렌더링 개선 내역 (ACS 저널 표준 최적화)

기본 Indigo 렌더러의 한계를 극복하고, 화학자 및 저널(ACS, IUPAC)이 선호하는 시각적 완성도를 갖추도록 다음과 같은 개선 로직이 적용되어 있습니다:

### ① 가변 고해상도 벡터급 렌더링 (Dynamic Resolution)
* 과거의 고정 해상도(예: 900×600) 강제 캔버스 방식을 탈피했습니다.
* **결합 길이(Bond Length: 80px)**, **여백(Margins: 30px)**, **선 상대 두께(Relative Thickness: 1.5)** 기반의 가변 해상도 렌더링을 적용하여, 분자 크기와 무관하게 모든 구조가 동일한 선 굵기와 비율로 선명하게 표시됩니다.

### ② Kekulé 표기법 자동 전환 (`dearomatize`)
* 벤젠 고리 등에서 어색한 점선 원형(Aromatic circle) 대신, 유기화학 논문 표준인 교대 이중결합(Kekulé form)으로 명확히 표현합니다 (`aromaticity-model = generic`).

### ③ 고리 접합부/가교 키랄 수소(H) 선별 보존 (`ExposeRingJunctionChiralHydrogens`)
* 일반 골격식에서는 탄소의 수소(H)를 숨기지만, 다환계 접합부(Ring Junction, 예: 아플라톡신, 베르게닌)에서는 결합선 겹침을 방지하고 입체 배치를 명확히 하기 위해 접합부 키랄 수소를 선별적으로 보존합니다.
* 동시에 α-투존(α-Thujone) 같은 3원환 내부의 비좁은 고리 결합에 쐐기선이 겹치는 시각적 왜곡 현상을 방지합니다.

### ④ 헤테로원자 결합 메틸기 및 소형 분자 가독성 개선
* **소형 분자 자동 분기:** 중원자 수 4개 이하의 초소형 분자(MIC `CN=C=O`, 메탄올 `CO`, 에탄올, 아세트산 등)는 빗금 하나(`\`) 대신 `H₃C-N=C=O`, `H₃C-OH` 등 화학 표준 텍스트로 명시합니다.
* **헤테로 결합 메틸기(`N-Me`, `O-Me`, `S-Me`) 자동 탐지:** 니코틴, 카페인, 아플라톡신, 베르게닌 등에서 N/O/S에 결합된 메틸기를 자동으로 식별하여 `N-CH₃` / `H₃C-O-` 등으로 표기합니다. 카로틴, 페로몬 등 순수 탄화수소 사슬의 메틸기는 골격식 선으로 유지되어 난잡함을 방지합니다.

### ⑤ 스마트 수평 레이아웃 (`smart-layout` + `layout-orientation: horizontal`)
* `smart-layout`을 활성화하여 티아민, 니코틴처럼 두 고리가 연결된 구조의 가교 결합을 반듯한 수평으로 정렬합니다.
* 가로가 넓은 엑셀 셀 종횡비에 최적화되도록 가로 우선 배치를 기본값으로 적용합니다.

### ⑥ 황(S) 원자 시인성 최적화
* 흰 배경에서 눈이 부시고 안 보이는 형광 노랑(`#FFFF00`) 대신, 흰 배경 대비를 고려한 다크 골든로드/황갈색(`#A88013`)으로 렌더링되어 높은 가독성을 제공합니다.

### ⑦ 투명 배경 및 사용자 정의 배경색 지원
* 셀 채우기 색상(노란색, 연회색 등)이 자연스럽게 비치는 투명 배경(Alpha=0)과 커스텀 배경색을 완벽히 지원합니다.

---

## 3. 함수 구문 및 인자 규격 (Function Specification)

```excel
=SMILES2IMG(smiles, [background], [color], [transform])
```

### 인자 상세 안내 (Parameters)

| 인자 순서 | 인자명 | 타입 | 필수 여부 | 기본값 | 설명 및 허용 값 |
| :---: | :--- | :---: | :---: | :---: | :--- |
| **1** | **`smiles`** | 문자열 | **필수** | - | 1~2000자의 유효한 SMILES 문자열 또는 셀 참조 (예: `A2`, `"CCO"`) |
| **2** | **`background`** | 문자열 | 선택 | `"white"` | **배경 색상 설정:**<br>• `"trans"` 또는 `"transparent"`, `"nobg"`: 투명 배경 (셀 배경 투과)<br>• `"white"`: 흰색 불투명 배경 (기본값)<br>• `"#RRGGBB"` 또는 `"bg=#RRGGBB"`: 사용자 지정 색상 (예: `"#FFFF00"`) |
| **3** | **`color`** | 불리언/숫자 | 선택 | `TRUE` | **원소 색상 모드 (학술 논문/인쇄용):**<br>• `TRUE` 또는 `1`: 산소(빨강), 질소(파랑), 황(황갈색) 등 원소별 표준 컬러 (기본값)<br>• `FALSE` 또는 `0`: 흑백(Grayscale) 모드 |
| **4** | **`transform`** | 문자열 | 선택 | `""` | **기하학적 변환 (회전 및 반전):**<br>• `flip`: 좌우 반전 (교과서/논문 표준 배치로 전환)<br>• `flipy`: 상하 반전<br>• `rot90`, `rot180`, `rot270`: 시계 방향 각도 회전 |

> [!TIP]
> **스마트 하위 호환성:** 2번째 인자 자리에 실수로 불리언 `FALSE`를 전달하더라도 파서가 자동으로 흑백 모드로 안전하게 인식합니다.

---

### 실전 활용 예제 (Practical Usage Scenarios)

#### 1. 기본 컬러 렌더링 (Default Color)
가장 일반적인 사용법입니다. 셀 참조 또는 화학식 문자열을 직접 전달합니다.
```excel
=SMILES2IMG(A2)
=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O")    ' 아스피린 직접 입력
```

#### 2. 가장 많이 쓰이는 투명 배경 (Transparent Background)
엑셀 시트에 표 서식(줄무늬), 조건부 서식, 또는 셀 채우기 색상(노랑, 연회색 등)이 적용되어 있을 때 흰색 박스 없이 자연스럽게 녹아듭니다.  
**쉼표를 건너뛸 필요 없이 2번째 인자에 바로 지정합니다.**
```excel
=SMILES2IMG(A2, "trans")                ' 컬러 + 투명 배경
=SMILES2IMG(A2, "transparent")          ' 풀 네임 사용 가능
```

#### 3. 사용자 지정 배경색 (Custom Hex Color)
특정 테마 색상으로 배경을 채우고 싶을 때 16진수 색상 코드를 2번째 인자에 전달합니다.
```excel
=SMILES2IMG(A2, "#FFFF00")              ' 노란색 배경
=SMILES2IMG(A2, "#FFF8DC")              ' 크림색 배경
```

#### 4. 학술 논문 / 인쇄용 흑백 표기 (ACS Style Black & White)
3번째 `color` 인자에 `FALSE` 또는 `0`을 지정하여 흑백 그레이스케일로 렌더링합니다.
```excel
=SMILES2IMG(A2, , FALSE)                ' 기본 흰 배경 + 흑백
=SMILES2IMG(A2, "trans", FALSE)         ' 투명 배경 + 흑백 (스타일끼리 나란히 지정)
```

#### 5. 교과서/화학 DB 표준 도식에 맞춘 좌우 반전 (`flip`)
SMILES의 원자 인덱싱 순서로 인해 고리 위치가 교과서나 화학 데이터베이스 도식과 반대로 놓이는 경우(예: 니코틴의 피롤리딘 고리, 티아민의 티아졸 고리 등), 4번째 `transform` 인자로 좌우를 뒤집습니다.
```excel
=SMILES2IMG(A2, , , "flip")             ' 기본 배경 + 좌우 반전
=SMILES2IMG(A2, "trans", , "flip")      ' 투명 배경 + 좌우 반전
=SMILES2IMG(A2, "trans", FALSE, "flip") ' 투명 배경 + 흑백 + 좌우 반전
```

#### 6. 90° / 180° / 270° 회전 및 상하 반전 (`rot90`, `flipy`)
긴 탄화수소 사슬이나 비대칭 분자가 엑셀 셀의 가로/세로 비율에 맞지 않을 때 회전시켜 가독성을 높입니다.
```excel
=SMILES2IMG(A2, , , "rot90")            ' 시계 방향 90도 회전
=SMILES2IMG(A2, "trans", , "rot90")     ' 투명 배경 + 90도 회전
=SMILES2IMG(A2, , , "flipy")            ' 상하 반전
```

---

### 실시간 수식 인텔리센스 (Formula IntelliSense) 안내

추가 기능 내에 `ExcelDna.IntelliSense`가 자체 내장되어 있어, 엑셀 기본 함수와 동일하게 셀 바로 아래에 실시간 안내 툴팁이 팝업됩니다:

```text
=SMILES2IMG(
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ SMILES2IMG(smiles, [background], [color], [transform])                                 │
│ A SMILES string or a cell reference containing one (e.g., A2, "CCO").                  │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

* **인자 이동:** 쉼표(`,`)를 누를 때마다 현재 입력 중인 인자가 **굵은 글씨**로 강조되며, 해당 인자의 영문 설명 툴팁이 자동으로 전환됩니다.
* **필요 없는 인자 생략:** 중간 인자를 기본값으로 유지하려면 `, ,` 처럼 연속 쉼표를 입력하여 다음 인자로 이동합니다 (예: `=SMILES2IMG(A2, , , "flip")`).

---

## 4. 빌드 및 배포 방법

### 사전 요구사항
* Windows OS
* [.NET SDK 8.0 이상](https://dotnet.microsoft.com/) (.NET Framework 4.8 타깃팅)

### 빌드 명령어
`cs_SMILES2IMG` 폴더에서 실행:

```powershell
dotnet restore
dotnet build -c Release --no-restore --tl:off
```

### 배포 산출물 구조 (`dist/`)
빌드가 완료되면 프로젝트 내 `dist/` 폴더에 32비트와 64비트 단일 파일 XLL 배포물이 생성됩니다:

```text
dist/
├── x64/
│   ├── Smiles2Img-AddIn64-packed.xll   # 64비트 Excel용 완전 독립형 단일 XLL (~8.6 MB)
│   └── THIRD_PARTY_NOTICES.md
└── x86/
    ├── Smiles2Img-AddIn-packed.xll     # 32비트 Excel용 완전 독립형 단일 XLL (~8.7 MB)
    └── THIRD_PARTY_NOTICES.md
```

> [!NOTE]
> **완전 자립형 단일 파일(Single-File) 배포:**  
> EPAM Indigo C++ 네이티브 렌더링 라이브러리 및 ExcelDna.IntelliSense가 `.xll` 파일 하나 안에 100% 압축 내장되어 있습니다.  
> 다른 DLL들을 번거롭게 함께 복사할 필요 없이, **`.xll` 파일 단 1개만** 원하는 위치에 두고 엑셀에 등록하면 즉시 정상 작동합니다 (첫 실행 시 `%LOCALAPPDATA%\Smiles2Img\native\`에 필요한 네이티브 바이너리가 1회 자동 추출됨).

---

## 5. 자동화 테스트 및 검증

### 단위 테스트 (Excel 불필요)
```powershell
dotnet build tests/Smiles2Img.Tests.csproj -c Release --no-restore
$env:PATH = "dist/x64;$env:PATH"
& tests/bin/Release/net48/Smiles2Img.Tests.exe
```
* 에탄올, 벤젠, 아스피린, 카페인 등 주요 분자의 PNG 렌더링 검증
* 소형 분자(`CN=C=O`) `H₃C-` 라벨링 및 `flip` 2D 변환 검증
* 투명 배경 알파 채널(Alpha=0) 및 커스텀 배경색 픽셀 무결성 검증
* 10개 대표 표준 분자 일괄 렌더링 검증

### Excel COM 자동화 스모크 테스트 (실제 Excel 구동)
```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests/Excel-Smoke.ps1
```
* 실제 Excel 프로세스를 백그라운드로 띄워 XLL 등록, 수식 계산, 인셀 이미지 삽입, 캐시 재사용, 통합문서 저장/재오픈, 수식 삭제 시 이미지 정리까지 전체 라이프사이클을 End-to-End로 검증합니다.
