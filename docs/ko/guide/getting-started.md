# 1분 설치 및 첫 분자 그리기

단 10초 만에 **`func_SMILES2IMG`**를 설치하고, Excel · Word · PowerPoint 화면에 첫 번째 화학 구조식을 띄워보세요.

<QuickInstallBox :hide-quickstart-btn="true" />

---

## Step 2. 오피스 재시작 & 상단 `[SMILES]` 리본 탭 확인

1. 열려 있던 Excel, Word, PowerPoint를 **모두 닫고 다시 실행**합니다.
2. 상단 메뉴 표시줄에 **`[SMILES]`** 탭이 생겼는지 확인하세요:

<p align="center">
  <img src="/excel-ribbon-guide.png" alt="오피스 상단 SMILES 리본 메뉴 탭" width="100%" style="border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);" />
</p>

---

## Step 3. 첫 분자 구조식 띄워보기 (내 프로그램 선택)

<p align="center">
  <img src="/hero-demo.gif" alt="func_SMILES2IMG 1분 빠른 시작 데모" width="100%" style="max-width: 760px; border-radius: 10px; box-shadow: 0 6px 20px rgba(0,0,0,0.12);" />
</p>

### 📗 A. 엑셀(Excel) 사용자라면? — 수식 복사해 붙여넣기
아무 빈 셀에나 아래 수식을 복사해 붙여넣고 <kbd>Enter</kbd>를 눌러보세요 (행 높이를 조금 넓히면 구조식이 더 크게 보입니다):

```text
=NAME2IMG("Aspirin")
```
또는 SMILES 문자열로 그리기:
```text
=SMILES2IMG("CN1C=NC2=C1C(=O)N(C(=O)N2C)C")
```
> 💡 **구버전 엑셀(Excel 2016 / 2019 / 2021)을 쓰시나요?**  
> 함수 이름 뒤에 **`.FLOAT`**를 붙여 `=NAME2IMG.FLOAT("Aspirin")` 또는 `=SMILES2IMG.FLOAT(A2)`로 입력하면 셀 크기에 맞춰 움직이는 그림으로 똑같이 생성됩니다!

### 📘📙 B. 워드(Word) · 파워포인트(PPT) 사용자라면? — `Ctrl + Shift + K` 누르기
Word와 PowerPoint는 셀 수식을 치는 대신 **리본 버튼이나 단축키**로 바로 그립니다:
1. 상단 리본 메뉴에서 **`[SMILES]` ➔ `[Insert / Edit Structure]`** 버튼을 클릭하거나, 키보드에서 <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>를 누릅니다.
2. 내장된 **Ketcher 2D 화학 편집기**가 열리면 원하는 분자나 합성 반응식을 그립니다.
3. <kbd>Ctrl</kbd> + <kbd>Enter</kbd> (또는 우측 상단 **Apply** 버튼)를 누르면 본문이나 슬라이드에 고화질 구조식이 즉시 들어갑니다!
4. **언제든 제자리 수정:** 삽입된 분자 그림을 **마우스로 더블클릭**하면 편집기가 0.05초 만에 다시 열려 바로 고칠 수 있습니다.

---

::: details ❓ 설치 후 엑셀에 `#NAME?` 오류가 뜨거나 `[SMILES]` 탭이 안 보이나요? (30초 해결)
설치 직후 수식이 `#NAME?`으로 표시되거나 리본 탭이 보이지 않는다면 아래 3가지를 순서대로 확인하세요:

1. **오피스를 완전히 껐다가 다시 켰나요?**  
   추가 기능은 오피스가 시작될 때 로드됩니다. 열려 있는 모든 Excel / Word / PowerPoint 창을 완전히 닫고 다시 실행해 주세요.
2. **Excel 2016 / 2019 / 2021 버전을 사용 중이신가요?**  
   셀 내부 그림 함수(`=SMILES2IMG`)는 Microsoft 365 및 Excel 2024 이상 전용입니다. 구버전 엑셀에서는 뒤에 **`.FLOAT`**만 붙여서 입력하시면 즉시 작동합니다:
   ```text
   =NAME2IMG.FLOAT("Aspirin")
   ```
3. **Windows 11 '스마트 앱 컨트롤' 차단 해제 (10초 해결):**  
   Windows 11 실행 시 *"파일 형식 및 확장명이 일치하지 않습니다"* 경고가 뜨거나 로드되지 않는 경우:
   - Windows **설정** (<kbd>Win</kbd> + <kbd>I</kbd>) ➔ **개인 정보 및 보안** ➔ **Windows 보안** ➔ **앱 및 브라우저 컨트롤**로 이동합니다.
   - **스마트 앱 컨트롤 설정**을 클릭해 **'꺼짐(Off)'**으로 변경한 뒤 오피스를 다시 실행하세요. *(스마트 앱 컨트롤을 꺼도 **Microsoft Defender 백신 실시간 감시**는 100% 그대로 정상 작동합니다.)*
:::

---

## 🔄 업데이트 및 깔끔한 제거 방법

* **최신 버전으로 업그레이드:**
  - 오피스 상단 **`[SMILES]` ➔ `[About SMILES]` ➔ `[Check for Updates]`**를 누르거나, 상단의 1줄 PowerShell 설치 명령어를 다시 실행하면 최신 버전으로 자동 덮어쓰기됩니다.
* **깔끔하게 제거하기 (Uninstall):**
  - PowerShell에 아래 한 줄을 입력하면 추가 기능 등록 및 파일이 완전히 삭제됩니다:
    ```powershell
    irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex
    ```

---

## 📚 사내 보안망(오프라인) 수동 설치 및 지원 사양 (클릭하여 펼치기)

::: details 🏢 사내 방화벽으로 자동 다운로드가 막혀 있는 경우 (오프라인 ZIP 설치)
1. [GitHub Releases](https://github.com/naramdash/func_SMILES2IMG/releases/latest)에서 `func_SMILES2IMG-x64.zip` (32비트 오피스는 `func_SMILES2IMG-x86.zip`)을 다운로드합니다.
2. 다운로드한 `.zip` 파일을 우클릭 ➔ **속성** ➔ 하단 **[차단 해제(Unblock)]** 체크 ➔ **[확인]**을 누릅니다.
3. 압축을 풀고 폴더 안에서 PowerShell을 열어 `powershell -ExecutionPolicy Bypass -File install.ps1`을 실행하면 오프라인으로 즉시 등록됩니다.
:::

::: details 💻 지원 환경 및 시스템 요구사항
| 항목 | 지원 사양 | 비고 |
| :--- | :--- | :--- |
| **운영체제** | Windows 10 / Windows 11 | 32비트 및 64비트 모두 지원 |
| **Microsoft Office** | Office 2016 / 2019 / 2021 / 2024 / Microsoft 365 | 데스크톱 설치 버전 |
| **필수 런타임** | .NET Framework 4.8 & Edge WebView2 | Windows 10/11에 기본 내장되어 별도 설치 불필요 |
:::
