# SMILES2IMG — JavaScript (Office.js + SmilesDrawer)

> 📌 **일반 Excel 사용자를 위한 빠른 설치 및 사용법 안내는 [루트 README.md](../README.md)를 참고하세요.**  
> 본 문서는 Office.js 웹 런타임 기반 구현에 대한 개발 및 배포 문서입니다.

Windows 데스크톱 Microsoft 365 Excel에서 `=SMILES.IMG(A2)`로 SMILES 문자열을 셀 안의 분자 이미지로 변환하는 추가 기능입니다. 네임스페이스는 `SMILES`, 함수명은 `IMG`입니다. Excel 웹 버전과 Google Sheets 대응은 이후 별도 대상으로 개발할 예정입니다.

## 사용 방법

1. 이 폴더에서 `npm ci`와 `npm start`를 실행해 추가 기능을 Excel에 로드합니다.
2. A2 셀에 `CCO`를 넣고 B2 셀에 `=SMILES.IMG(A2)`를 입력합니다.
3. 이미지가 잘 보이도록 행 높이와 열 너비를 늘립니다.

`=SMILES.IMG("c1ccccc1")`처럼 문자열을 직접 입력해도 됩니다. 이미지는 기본 300 × 200 픽셀입니다. 기존 CONTOSO 네임스페이스로 추가 기능을 등록했다면 `npm stop` 후 Excel을 종료하고 `npm start`로 다시 등록합니다.

입력은 공백을 제거한 1~2000자의 SMILES 문자열입니다. 빈 입력·잘못된 SMILES는 `#VALUE!`, 이미지 생성 실패나 기본 요구사항을 충족하지 않는 Excel은 `#N/A`로 처리합니다. Excel 자체가 Preview 이미지 타입을 처리하지 못하면 결과가 `#VALUE!`로 표시될 수 있습니다.

## 로컬 이미지 생성과 Base64 반환

이미지 생성은 Excel 추가 기능의 웹 런타임 안에서 이루어집니다. 번들에 포함된 SmilesDrawer로 SMILES를 해석하고 300 × 200 픽셀 Canvas에 분자 구조를 그린 뒤, Canvas를 PNG Base64로 인코딩해 `Excel.LocalImageCellValue`를 반환합니다.

```text
=SMILES.IMG(A2)
  → 추가 기능 런타임에서 SmilesDrawer로 SMILES 해석
  → Canvas → PNG Base64
  → { type: "LocalImage", image: { type: "PNG", data: "..." } }
  → Excel이 이미지 데이터를 셀에 표시
```

`image.data`에는 `data:image/png;base64,` 접두어를 제외한 순수 Base64 문자열을 넣습니다. 이미지 URL이나 이미지 생성 API를 사용하지 않습니다. SMILES는 이미지 생성을 위해 외부 서버로 전송되지 않습니다.

`LocalImageCellValue`와 `Base64EncodedImage`는 현재 `ExcelApi BETA`의 Preview API입니다. 함수 실행용 페이지는 Office.js Preview CDN을 사용하고, TypeScript는 `@types/office-js-preview`를 사용합니다. 지원되는 Excel 빌드에서 셀 이미지 표시를 확인해야 합니다. [로컬 이미지 API](https://learn.microsoft.com/en-us/javascript/api/excel/excel.localimagecellvalue?view=excel-js-preview), [Base64 이미지 스키마](https://learn.microsoft.com/en-us/javascript/api/excel/excel.base64encodedimage?view=excel-js-preview), [Preview 사용 방법](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-preview-apis)

작업창의 **분자 미리보기**에서 이미지를 확인하고 **PNG 저장**을 사용할 수 있습니다. 미리보기는 일반 브라우저에서도 확인할 수 있지만, 워크시트의 `SMILES.IMG` 실행에는 Excel이 필요합니다.

## 서버가 필요한 범위

이미지 생성용 Node.js 서버, `server/image-service.cjs`, 이미지 생성 API, `sharp`는 제거했습니다.

개발 중의 Webpack HTTPS 서버는 추가 기능 HTML·JS 파일을 제공하는 역할만 합니다. 배포 시에는 `dist/`를 HTTPS 정적 호스팅에 올릴 수 있습니다. 사용자 PC에서 Node.js 이미지 서버를 실행할 필요는 없습니다. [Office 추가 기능의 호스팅 구조](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

SmilesDrawer는 추가 기능 JS 번들에 포함되며 별도 스크립트나 WASM 파일을 내려받지 않습니다. 추가 기능을 불러온 뒤 새 분자의 PNG를 생성할 때 이미지 생성 서버나 네트워크 요청이 필요하지 않습니다. Excel을 다시 열거나 추가 기능을 다시 로드하는 경우 파일 제공 주소에 다시 접근할 수 있어야 하므로 완전한 오프라인 추가 기능을 보장하는 구성은 아닙니다.

## 주요 파일

| 파일 | 역할 |
| --- | --- |
| `src/functions/functions.ts` | Excel 사용자 지정 함수 정의 |
| `src/commands/commands.ts` | Excel 리본 명령 초기화와 `action` 핸들러 기본 틀 |
| `src/taskpane/` | 추가 기능 작업창의 HTML, CSS, TypeScript |
| `src/rendering/molecule.ts` | SmilesDrawer 기반 로컬 PNG Base64 생성과 캐시 |
| `src/runtime.ts` | Office.js Preview 라이브러리와 데이터 타입 반환 기본 요구사항 확인 |
| `tests/` | 함수 반환·오류·캐시·재시도, 파서·번들·메타데이터 검증 |
| `manifest.xml` | 함수 네임스페이스, 코드·메타데이터 URL, Shared Runtime, 권한 설정 |
| `webpack.config.js` | SmilesDrawer 번들링과 함수 메타데이터 생성 |
| `babel.config.json` | TypeScript 문법 제거와 브라우저용 JavaScript 변환 |
| `tsconfig.json` | TypeScript 타입 검사 설정 |

## 개발 환경과 실행

Node.js 24.x에서 24.11.0 이상을 권장합니다. 현재 검증한 환경은 Node.js 24.17.0입니다. Babel 8의 Node.js 요구사항을 만족해야 합니다.

Excel은 LocalImage Preview를 지원하는 Windows Microsoft 365 데스크톱 빌드가 필요합니다. 최신 Microsoft 365 Excel 또는 Insider 빌드에서 확인합니다. LocalImage Preview의 공식 최소 지원 빌드는 공개되어 있지 않습니다.

매니페스트의 `SharedRuntime 1.1`, `CustomFunctionsRuntime 1.4`, `ExcelApi 1.16`은 안정 API의 기본 요구사항을 지정합니다. 이 조건을 충족해도 LocalImage Preview 지원이 보장되지는 않습니다. 런타임에서는 `CustomFunctionsRuntime 1.4`와 `Excel.CellValueType.localImage` 존재 여부를 검사하며, Preview 열거형의 존재만으로 Excel 내부의 실제 지원 여부를 판별할 수는 없습니다.

Excel 2019와 영구 라이선스 Excel은 현재 구현의 대상이 아닙니다. [사용자 지정 함수 요구사항](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)

아래 명령은 이 README가 있는 `SMILES2IMG` 폴더에서 실행합니다.

```powershell
# 잠금 파일 기준 의존성 설치
npm ci

# 타입 검사: JavaScript 파일을 생성하지 않음
npx tsc

# 함수·SmilesDrawer·메타데이터·번들 테스트
npm test

# 프로덕션 빌드 / 개발 빌드
npm run build
npm run build:dev

# 매니페스트 검증
npm run validate

# 데스크톱 Excel에 추가 기능을 로드하고 개발 서버 시작
npm start

# 개발 서버 종료와 개발용 추가 기능 등록 해제
npm stop
```

`npm start`는 `prestart`를 통해 빌드를 먼저 실행합니다. 개발 서버 주소는 `https://localhost:3000/`입니다. 최초 실행에서는 개발 인증서와 WebView2의 localhost 접근 설정이 필요할 수 있습니다. 자세한 절차는 [사용자 지정 함수 빠른 시작](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-custom-functions-quickstart?tabs=excel-windows)을 참고합니다.

## 사용자 지정 함수를 만드는 방법

### 함수 정의와 메타데이터

`src/functions/functions.ts`에 함수를 내보내고, JSDoc에 `@customfunction`을 붙입니다. 아래 ADD 코드는 등록 원리를 설명하는 예시입니다. 현재 프로젝트에는 `IMG`만 등록됩니다.

```typescript
/**
 * 두 숫자를 더합니다.
 * @customfunction ADD
 * @param first 첫 번째 숫자.
 * @param second 두 번째 숫자.
 * @returns 두 숫자의 합.
 */
export function add(first: number, second: number): number {
  return first + second;
}
```

TypeScript의 함수 선언으로 인수·반환 타입을 지정합니다. 설명과 `@param`, `@returns`는 Excel의 함수 도움말에 반영됩니다. `@customfunction`에 ID를 생략하면 함수 이름에서 ID가 생성되므로, 공개 함수에는 명시적인 ID를 사용하는 것이 좋습니다.

빌드 시 `CustomFunctionsMetadataPlugin`이 함수 선언을 읽어 `functions.json`을 생성합니다. 이 메타데이터에는 함수 ID, 이름, 인수, 반환 타입 등이 기록됩니다. 일반 빌드 결과는 `dist/`에 생성되며, 개발 서버는 결과물을 메모리에서 제공할 수 있습니다. 생성된 메타데이터를 직접 수정하기보다 함수 소스와 빌드 설정을 수정합니다. 함수 정의를 여러 파일로 분리하면 플러그인의 `input`과 Webpack의 함수 진입점도 함께 갱신해야 합니다. [메타데이터 자동 생성 문서](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json-autogeneration)

### 함수 이름과 네임스페이스

현재 `manifest.xml`의 `Functions.Namespace` 값은 `SMILES`이고, 함수 소스의 `@customfunction IMG`가 함수 ID와 이름을 지정합니다.

```excel
=SMILES.IMG(A2)
```

빌드 플러그인이 `CustomFunctions.associate("IMG", img)`를 자동으로 추가합니다. `functions.json`의 `allowCustomDataForDataTypeAny: true`가 이미지 객체 반환을 허용합니다. 메타데이터 생성기는 결과 타입 `any`를 기본값으로 보고 `result: {}`로 출력합니다. [함수 이름과 네임스페이스 문서](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-naming), [사용자 지정 함수의 데이터 타입](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-data-types-concepts)

### 비동기 처리와 오류

이미지 디코딩이나 웹 요청을 기다려야 하는 함수는 `async` 함수로 작성하고 `Promise`로 결과를 반환합니다. Excel은 해당 결과가 준비되기를 기다립니다. 웹 요청에서는 `fetch()`가 완료되었더라도 `response.ok`를 별도로 확인해야 합니다.

실패는 `CustomFunctions.Error`와 적절한 `CustomFunctions.ErrorCode`로 전달합니다. 잘못된 입력과 이미지 생성·통신 실패를 구분해 사용자가 원인을 알 수 있도록 처리합니다. [튜토리얼의 웹 요청 예제](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions#create-a-custom-function-that-requests-data-from-the-web)

### 일회성 결과와 스트리밍

입력 하나에 결과 하나를 반환하는 SMILES 변환에는 일반 비동기 함수가 적합합니다. 반복적으로 셀을 갱신해야 하는 경우에는 `CustomFunctions.StreamingInvocation`과 `invocation.setResult()`를 사용합니다. 스트리밍이 취소될 때는 `onCanceled`에서 타이머, 연결 등의 자원을 정리합니다. [비동기·스트리밍 문서](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs)

### 재등록과 캐시

새 함수를 추가한 뒤에는 재빌드하고 Excel에서 추가 기능을 다시 등록합니다. 데스크톱에서는 Excel을 종료했다가 다시 열어 개발용 추가 기능을 등록하는 절차가 필요할 수 있습니다. 함수 이름·인수가 갱신되지 않거나 중복 등록 오류가 나면 `npm stop` 후 Office 캐시를 확인합니다. [튜토리얼의 등록·문제 해결 절차](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions)

## SmilesDrawer 렌더링과 배포

- `smiles-drawer@2.4.1`을 일반 JS 의존성으로 import하고 Webpack 번들에 포함합니다. 외부 렌더러 파일, WASM, 이미지 API 요청은 없습니다.
- 동일 SMILES의 동시 계산은 하나의 렌더링을 공유하고, 최근 128개 PNG 결과를 메모리에 캐시합니다.
- SmilesDrawer 2.x는 내부적으로 벡터 도형을 만든 뒤 Canvas로 변환합니다. 로컬 이미지의 `decode()` 완료를 기다려 빈 Canvas가 반환되지 않도록 합니다. 사용자에게 반환하거나 저장하는 형식은 PNG뿐입니다.
- 파싱 오류는 `#VALUE!`, 이미지 디코딩·렌더링·인코딩 오류는 `#N/A`로 전달합니다. 실패한 결과는 캐시하지 않아 다음 호출에서 재시도할 수 있습니다.
- SmilesDrawer는 SMILES 파서와 구조 렌더러입니다. RDKit의 화학 구조 검증이나 분석 기능을 제공하는 것으로 가정하지 않습니다.

배포할 때는 `dist/`의 모든 파일을 HTTPS 정적 호스팅에 함께 올립니다. `ADDIN_BASE_URL` 환경 변수로 프로덕션 매니페스트의 localhost URL을 배포 주소로 치환할 수 있습니다. 지정하지 않으면 로컬 개발 주소를 유지합니다. 주소가 하위 경로를 포함한다면 마지막에 `/`를 붙입니다.

사용한 API는 [SmilesDrawer 공식 저장소](https://github.com/reymond-group/smilesDrawer)를 참고합니다.

## 현재 빌드 설정의 기준

- `module: ESNext`, `moduleResolution: Bundler`: Webpack으로 묶는 브라우저 코드에 맞춘 모듈 설정입니다. `node16` 등의 값은 Node.js 실행 버전 지정이 아니라 모듈 처리 규칙입니다.
- `target`과 `lib`: `ES2025`, `DOM`, `DOM.Iterable`을 사용합니다. 실제 브라우저용 문법 변환은 Babel이 담당합니다.
- `noEmit: true`: TypeScript는 타입 검사에 사용하고 JavaScript 생성은 Babel·Webpack이 담당합니다. Webpack 빌드가 타입 검사까지 수행하는 것은 아니므로 `npx tsc`도 별도로 실행합니다.
- `isolatedModules: true`: Babel처럼 파일 단위로 변환하는 도구와 맞지 않는 TypeScript 사용을 검사합니다.
- `browserslist: ["last 2 Edge versions"]`: Windows Excel의 Edge WebView2를 대상으로 합니다. `last 2 versions`만 쓰면 IE 등 다른 브라우저도 포함되므로 사용하지 않습니다. Mac Excel 지원을 추가할 때는 Safari 타깃도 검토합니다.
- TypeScript는 `6.0.3`을 사용합니다. 함수 메타데이터 생성기가 사용하는 컴파일러 API를 TypeScript 7.0이 제공하지 않아 직접 교체하지 않았습니다.
- lint와 Prettier 관련 패키지·스크립트·설정은 제거했습니다. `core-js/stable`과 `regenerator-runtime/runtime`은 현재 번들에 포함되어 있습니다.

설정의 근거는 [TypeScript 번들러 설정 가이드](https://www.typescriptlang.org/docs/handbook/modules/guides/choosing-compiler-options.html), [Office의 브라우저 엔진 문서](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins), [TypeScript 7의 컴파일러 API 설명](https://devblogs.microsoft.com/typescript/announcing-typescript-7-0/#running-side-by-side-with-typescript-6-0)을 참고합니다.

자동 테스트는 LocalImage 반환 스키마, 입력 오류, 디코딩 완료 대기·재시도, 이미지 캐시, 실제 SmilesDrawer SMILES 해석, 함수 메타데이터와 번들 제공을 검증합니다. 함수 단위 테스트의 Canvas는 테스트 대역을 사용합니다.

SmilesDrawer로 교체한 뒤 실제 Edge 브라우저에서 PNG Base64 생성·미리보기와 추가 기능 로딩 후 네트워크를 끊고 새 분자를 그리는 동작을 확인했습니다. Excel 앱 안에서의 최종 LocalImage 셀 표시는 아직 확인하지 않았습니다.
