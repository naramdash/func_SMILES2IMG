# WebView2 로딩 오류

내장 **Ketcher 2D** 편집기 또는 **Mol\* 3D** 단백질 뷰어 창이 빈 화면으로 열리거나 초기화되지 않을 때의 해결 방법입니다.

## 1. Microsoft Edge WebView2 런타임 확인

Ketcher 2D와 Mol\* 3D는 모두 **Microsoft Edge WebView2** 런타임 위에서 로컬로 동작합니다.

* Windows 10/11에 기본 탑재되어 있으나 일부 경량화 사내 Windows 이미지에서는 누락되어 있을 수 있습니다.
* WebView2 초기화 오류가 발생하면 공식 **Microsoft Edge WebView2 Evergreen Runtime**을 설치한 뒤 오피스를 재시작하세요.

<FeatureMedia id="common-ketcher-2d" />

## 2. 로컬 에디터 번들 캐시 초기화

`func_SMILES2IMG`는 오프라인 Ketcher/Mol\* 자산을 `%LOCALAPPDATA%\Smiles2Img`에 추출하고 `SHA-256` 지문으로 무결성을 검증합니다. 백신 잠금이나 비정상 종료로 파일이 손상된 경우:

1. 열려 있는 Excel, Word, PowerPoint를 모두 종료합니다.
2. <kbd>Win</kbd> + <kbd>R</kbd> ➔ `%LOCALAPPDATA%\Smiles2Img`를 입력해 폴더를 엽니다.
3. 폴더 안의 추출된 캐시 하위 폴더를 삭제합니다.
4. 오피스를 다시 열고 Ketcher 또는 Mol\*을 실행하면 깨끗한 번들이 자동으로 다시 추출됩니다.

## 3. 고해상도 DPI 배율 또는 창이 화면 밖에 열린 경우

편집기 창은 모니터 DPI(`100%`~`200%`)에 맞춰 자동 조절됩니다. 보조 모니터를 분리한 직후 창이 화면 밖에 걸렸다면 <kbd>Alt</kbd> + <kbd>Space</kbd> ➔ **M** 후 방향키로 이동하거나, 창을 닫고 `[SMILES]` 리본에서 다시 여세요.

## 관련 가이드

- [**오프라인 PDB 파일**](/ko/editor/troubleshooting/offline-pdb) — 사내 방화벽 환경에서 로컬 .pdb, .cif, .sdf 파일 열기
- [**Ketcher 2D 편집기 가이드**](/ko/editor/) — 단축키, 다단계 반응식 그리기, 배경 모드 사용법
