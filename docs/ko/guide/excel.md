# SMILES for Microsoft Excel 가이드

**SMILES for Microsoft Excel**은 엑셀 워크시트 셀 안에 2D 화학 구조식을 네이티브 인셀 이미지 또는 플로팅 도형으로 실시간 렌더링하고, 셀 더블 클릭으로 제자리 수정을 지원하는 화학 연구용 추가 기능입니다.

---

## 🎥 3분 핵심 기능 마스터 영상

> 💡 **SMILES for Excel의 모든 핵심 기능을 1분 56초 만에 마스터할 수 있는 1080p 고해상도 가이드 영상입니다.**  
> 인셀 수식 입력, 전문 화학 스타일링, 오프라인 Ketcher 에디터 연동, 화학 반응식(`=KET2IMG`), 원클릭 물성 정보 복사까지 한눈에 확인하세요.

<p align="center">
  <video controls="controls" width="100%" style="max-width: 800px; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.15);" poster="/excel-showcase.png">
    <source src="/videos/excel-guide-video.mp4" type="video/mp4" />
    <source src="https://github.com/user-attachments/assets/e566a830-55d5-4ff5-8030-344d8b8ce7b6" type="video/mp4" />
    <a href="/videos/excel-guide-video.mp4">🎬 Excel 핵심 기능 3분 마스터 동영상 재생하기</a>
  </video>
</p>

<details>
<summary><b>⏱️ 영상 타임라인 및 챕터별 내용</b></summary>

| 시간 | 챕터 및 주제 | 핵심 내용 및 자막 요약 |
| :---: | :--- | :--- |
| **00:00** | **개요 및 핵심 소개** | 네이티브 인셀 이미지 특징, 100% 로컬 오프라인 구동 및 셀 크기 자동 연동 |
| **00:13** | **1단계: 기본 인셀 수식** | 빈 셀에 `=SMILES2IMG(A2)` 수식 입력 및 수백 개 화합물 일괄 드래그 채우기 |
| **00:28** | **2단계: 전문 스타일 옵션** | 특허청 선화(`bw`), NMR 피크 번호(`num`), 입체화학(`stereo`), 다크모드 및 회전/반전 |
| **00:44** | **3단계: Ketcher 2D 에디터** | 단축키 `Ctrl+Shift+K` 또는 우클릭으로 오프라인 에디터 실행 및 더블 클릭 제자리 수정 |
| **01:00** | **4단계: 화학 반응식 (KET)** | 반응 화살표와 텍스트 주석이 보존되는 `=KET2IMG(A2)` 캔버스 수식 활용법 |
| **01:15** | **5단계: 퀵 인서트 & 물성 복사** | 클립보드 즉시 삽입 및 화학식(Formula)·분자량(MW, g/mol) 원클릭 계산 |
| **01:30** | **핵심 요약 & 체크리스트** | 엑셀 실전 업무 및 논문 준비를 위한 4대 핵심 체크포인트 정리 |

</details>

---

## 🌟 한눈에 보는 핵심 기능

<p align="center">
  <img src="/excel-showcase.png" alt="SMILES for Excel Showcase" width="100%" />
</p>

::: tip ✨ 4줄 요약: 이것만 알면 끝!
1. **빈 셀에 바로 그리기:** <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd> 누르고 그려서 <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (셀에 적용).
2. **수식으로 대량 삽입:** `=SMILES2IMG(A2)` 또는 `=KET2IMG(A2)` 입력 후 아래로 드래그.
3. **클립보드 즉시 삽입:** SMILES/KET 복사 후 리본 메뉴의 **[클립보드에서 붙여넣기]** 클릭.
4. **언제든 수정/정보확인:** 셀 더블 클릭으로 제자리 수정 / 우클릭 **[Chemical Structure]**로 분자 정보 즉시 복사.
:::

---

## 📝 수식 함수: 네이티브 인셀 vs. 플로팅 이미지

엑셀 버전에 따라 최적의 수식 함수를 선택할 수 있습니다:

### 1. `=SMILES2IMG(smiles_or_ket, [bg], [options])`
* **지원 버전:** **Microsoft 365** 및 **Excel 2024 이상**
* 셀 내부에 진짜 **인셀(In-Cell) 그림**으로 쏙 들어갑니다.
* 행 높이와 열 너비를 조절하면 분자 그림이 비율을 유지하며 자동으로 크기가 맞춰집니다.
* 정렬, 필터, 셀 복사/이동 시 셀 데이터와 일체형으로 함께 동작합니다.

### 2. `=SMILES2IMG.FLOAT(smiles_or_ket, [bg], [options])`
* **지원 버전:** **Excel 2016, 2019, 2021** (또는 셀 위 자유 배치를 원하는 경우)
* 셀 좌표에 고정(`xlMoveAndSize`)되는 플로팅 그림 도형을 생성합니다.
* 수식이 바뀌면 기존 도형을 스스로 찾아 깨끗이 치환하여 도형이 중복으로 쌓이지 않습니다.

---

## 🎨 화학 스타일 옵션

옵션 인수에 띄어쓰기, 쉼표, 세미콜론 등으로 자유롭게 조합할 수 있습니다:

| 옵션 키워드 | 설명 | 수식 예시 |
| :--- | :--- | :--- |
| `bw` / `monochrome` | 특허청/ACS 논문용 100% 흑백 선화 모드 | `=SMILES2IMG(A2, "", "bw")` |
| `colored` | 원소별 표준 색상 강조 모드 (기본값) | `=SMILES2IMG(A2, "", "colored")` |
| `num` / `numbers` | 원자 인덱스 번호 표기 (NMR 피크 매핑에 필수) | `=SMILES2IMG(A2, "", "num")` |
| `stereo` | 입체화학 라벨(*R/S*, *E/Z*) 강제 표시 | `=SMILES2IMG(A2, "", "stereo")` |
| `d=2.0` / `d=1.5` | 고해상도 슈퍼샘플링 배율 (초고선명 인쇄용) | `=SMILES2IMG(A2, "", "d=2.0")` |
| `flip` / `flipv` | 좌우 / 상하 구조식 반전 | `=SMILES2IMG(A2, "", "flip")` |
| `rot90` / `rot180` | 90도 / 180도 분자 회전 | `=SMILES2IMG(A2, "", "rot90")` |
| `w=300 h=200` | 플로팅 이미지용 픽셀 크기 지정 | `=SMILES2IMG.FLOAT(A2, "", "w=300 h=200")` |

<p align="center">
  <img src="/excel-matrix-showcase.png" alt="Excel Matrix Styling Showcase" width="100%" />
</p>

---

## 🧪 무손실 화학 반응식: `=KET2IMG`

반응 화살표, 반응 조건 텍스트, 촉매, 플러스 기호가 포함된 다단계 합성 반응식:

```text
=KET2IMG(A2)
```
* **무손실 보존:** 반응식 전체 구조와 캔버스 좌표가 KET 포맷으로 셀 수식에 영구 보존됩니다.
* **제자리 수정:** 셀을 더블 클릭하면 반응 화살표와 텍스트까지 그대로 유지된 채 Ketcher에서 열려 즉시 수정할 수 있습니다.

---

## 📋 우클릭 메뉴 & 분자 정보 원클릭 복사

화학 구조식이 들어있는 셀이나 빈 셀을 우클릭하면 전용 메뉴가 표시됩니다:

* **Chemical Structure ➔ Ketcher로 구조 편집:** 오프라인 에디터 실행.
* **Chemical Structure ➔ SMILES / KET 복사:** 순수 구조식 문자열 클립보드 복사.
* **Chemical Structure ➔ 분자식(Formula) 복사:** Hill 체계 화학식 자동 계산 (예: `C8H10N4O2`).
* **Chemical Structure ➔ 분자량(MW) 복사:** 화학 분자량 자동 계산 (g/mol, 예: `194.19`).
* **Chemical Structure ➔ 정밀 질량(Exact Mass) 복사:** 고분해능 단일동위원소 질량 산출.
