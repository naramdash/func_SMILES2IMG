# 1분 설치 및 첫 분자 그리기

관리자 권한 없이 10초 만에 **`func_SMILES2IMG`**를 설치하고, 사용 중인 오피스 앱에 맞는 전용 가이드(**Excel 레시피 북**, **Word 6장 논문 튜토리얼**, **PowerPoint 슬라이드 갤러리**)로 바로 시작하세요.

---

## Step 1. 10초 PowerShell 자동 설치

<QuickInstallBox />

::: tip 🚀 `v1.8.0` 정식 포함 기능
위 명령어로 설치되는 **`func_SMILES2IMG v1.8.0`**에는 Excel·Word·PowerPoint 통합 `Chemical Structure >` 우클릭 메뉴, 앱 간 물질명·프리로드·비율 고정 통일, 2D 미리보기 표준 Windows Molecule Info/About 창, `=SMILES2HEAVY`~`=SMILES2CHIRAL` 3D 복잡도 함수 5종, 12열 `.sdf` 가져오기가 포함되어 있습니다.
:::

---

## Step 2. 오피스 재시작 & `[SMILES]` 리본 탭 확인

1. 열려 있던 Excel, Word, PowerPoint를 **모두 닫고 다시 실행**합니다.
2. 상단 메뉴 표시줄에 **`[SMILES]`** 탭이 나타났는지 확인하고, 언제든 **`[About func_SMILES2IMG]`**를 클릭해 버전과 업데이트를 확인할 수 있습니다.

<FeatureMedia id="common-about-dialog" />

---

## Step 3. 내가 쓰는 앱 선택 & 첫 구조식 그리기

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="./excel" class="s2i-card">
    <div class="s2i-card-title">📗 Microsoft Excel — 실무 레시피 북</div>
    <p class="s2i-card-desc">구조식 표, 이름/CAS 변환, Lipinski 약물성, SAR 스크리닝 레시피와 40개 함수 사전</p>
  </a>
  <a href="./word" class="s2i-card">
    <div class="s2i-card-title">📘 Microsoft Word — 6장 논문 튜토리얼</div>
    <p class="s2i-card-desc">실제 연구 논문 원고 1편을 서론 삽입부터 투고용 ACS 흑백 통일까지 완성해 나갑니다</p>
  </a>
  <a href="./powerpoint" class="s2i-card">
    <div class="s2i-card-title">📙 Microsoft PowerPoint — 슬라이드 갤러리</div>
    <p class="s2i-card-desc">다크 모드 투명 구조식, 반응 메커니즘, 360° 회전 3D 단백질 슬라이드를 보고 따라 합니다</p>
  </a>
</div>

### 📗 A. Microsoft Excel 사용자라면? (수식 입력)
아무 빈 셀에나 아래 수식을 붙여넣고 <kbd>Enter</kbd>를 눌러보세요 (Excel 2016~2021은 뒤에 **`.FLOAT`** 추가: `=NAME2IMG.FLOAT("Aspirin")`):

```text
=NAME2IMG("Aspirin")
```
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```

<FeatureMedia id="excel-overview" />

::: tip 📥 8개 탭 마스터 쇼케이스 엑셀 파일 다운로드
수식을 직접 치지 않아도 40종 전체 함수와 예제가 동작하는 **8개 탭 마스터 워크북**을 바로 열어볼 수 있습니다:
<div style="margin: 10px 0 4px 0;">
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card" style="display: block; padding: 12px 16px; border: 1.5px solid var(--vp-c-brand-1); border-radius: 8px; text-decoration: none;">
    <div style="font-weight: 700; color: var(--vp-c-brand-1); font-size: 1rem;">📥 SMILES2IMG_Comprehensive_Showcase.xlsx 다운로드 (.xlsx)</div>
    <div style="color: var(--vp-c-text-2); font-size: 0.88rem; margin-top: 4px;">0. 빠른 시작 · 1. 이름/CAS/SMILES · 2. Lipinski &amp; 3D 복잡도 · 3. 부분구조/유사도 · 4. 스타일 매트릭스 · 5. KET 반응식 · 6. 구형 엑셀(.FLOAT) · 7. 3D &amp; 리본 도구</div>
  </a>
</div>
:::

### 📘 B. Microsoft Word 사용자라면? (본문 · 표 삽입)
상단 **`[SMILES]` ➔ `[Insert / Edit Structure]`** (또는 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>)로 Ketcher 편집기를 열어 그린 뒤 <kbd>Ctrl</kbd> + <kbd>Enter</kbd>를 누르면 커서 위치에 인라인 그림으로 들어갑니다.

<FeatureMedia id="word-overview" />

### 📙 C. Microsoft PowerPoint 사용자라면? (다크 모드 & 3D 슬라이드)
다크 테마 슬라이드에 흰색 테두리 없이 투명 구조식을 배치하거나, 발표 중 마우스로 돌리는 3D 단백질을 삽입합니다.

<FeatureMedia id="powerpoint-overview" />

---

::: details ❓ 설치 후 `#NAME?` 오류가 뜨거나 `[SMILES]` 탭이 안 보이나요? (30초 해결)
1. **오피스를 완전히 껐다가 다시 켰나요?**  
   추가 기능은 오피스 시작 시 로드됩니다. 모든 창을 닫고 다시 실행해 주세요.
2. **Excel 2016 / 2019 / 2021을 사용 중이신가요?**  
   인셀 이미지 함수(`=SMILES2IMG`)는 Microsoft 365 및 Excel 2024 이상 전용입니다. 구버전에서는 **`.FLOAT`**를 붙이세요 (`=NAME2IMG.FLOAT("Aspirin")`).
3. **Windows 11 스마트 앱 컨트롤 차단 해제:**  
   **설정** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **개인 정보 및 보안** ➔ **Windows 보안** ➔ **앱 및 브라우저 컨트롤** ➔ **스마트 앱 컨트롤 설정**을 **꺼짐(Off)**으로 바꾼 뒤 오피스를 재시작하세요. *(Microsoft Defender 백신 실시간 감시는 그대로 유지됩니다.)*
:::

---

## 🔄 업데이트 및 깔끔한 제거

* **업데이트:** 리본 **`[SMILES]` ➔ `[About func_SMILES2IMG]` ➔ `[Check for Updates]`**를 누르거나, 상단 설치 명령어를 다시 실행하세요.
* **깔끔한 제거 (Uninstall):**
  ```powershell
  irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
  ```

---

## 📚 사내 보안망(폐쇄망) 수동 설치 & 시스템 사양

::: details 🏢 오프라인 ZIP 수동 설치 방법
1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest)에서 `func_SMILES2IMG-x64.zip` (32비트 오피스는 `x86`)을 내려받습니다.
2. 파일 우클릭 ➔ **속성** ➔ 하단 **[차단 해제]** 체크 ➔ **[확인]** 클릭.
3. 압축 해제 후 해당 폴더에서 PowerShell을 열고 `powershell -ExecutionPolicy Bypass -File install.ps1`을 실행합니다.
:::

::: details 💻 지원 환경 및 시스템 요구사항
| 항목 | 지원 사양 | 비고 |
| :--- | :--- | :--- |
| **운영체제** | Windows 10 / Windows 11 | 32비트 및 64비트 모두 지원 |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | 데스크톱 설치 버전 |
| **필수 런타임** | .NET Framework 4.8 & Edge WebView2 | Windows 10/11 기본 내장 |
:::
