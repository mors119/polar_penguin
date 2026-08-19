# Polar Penguin

Google Apps Script로 주문, 재고, 예약재고, 피킹을 연결하는 운영 시스템입니다.

사용자는 카페24에서 내려받은 주문 또는 재고 파일을 구분할 필요 없이 `Input` 폴더에 올리면 됩니다. 시스템이 헤더로 파일 유형을 판별하고 검증, 주문/재고 반영, 피킹 생성, PDF 출력, 원본 이동, 오류 알림과 대시보드 갱신을 수행합니다.

## 설치

### 1. Apps Script에 코드 반영

```bash
npm install
clasp push
```

`.clasp.json`은 기존 Apps Script 프로젝트의 Script ID를 유지하며, `src/` 안의 파일만 push합니다.

### 2. ROOT 폴더 지정

Apps Script 편집기에서 한 번 실행합니다.

```javascript
setRootFolder('https://drive.google.com/drive/folders/...');
```

또는 Apps Script의 프로젝트 설정 → Script Properties에 다음 중 하나를 직접 등록합니다.

```text
ROOT_FOLDER_ID=<Google Drive folder ID>
ROOT_FOLDER_URL=<Google Drive folder URL>
```

### 3. 자동 설치

```javascript
setupSystem();
```

첫 실행 시 Drive, Spreadsheet, Script trigger, Gmail 권한 요청을 승인합니다. `setupSystem()`은 다음 구조를 생성하거나 기존 리소스를 재사용합니다.

> Only one operational Google Spreadsheet is used.

```text
ROOT_FOLDER/
├── Polar Penguin
├── Input/
├── Processed/
├── Error/
├── Output/
└── Backup/
```

`Polar Penguin`은 하나의 Google Spreadsheet이며 다음 탭을 가집니다.

```text
📖 안내
📊 대시보드
상품마스터
주문(완료)
피킹(헤더)
피킹(라인)
예약대기
주문반려
재고이동로그
작업로그
입력처리로그
설정
```

`예약대기`와 `주문반려`는 `주문(완료)`의 현재 상태에서 갱신되는 조회용 탭입니다. 별도의 원장이 아닙니다. setup은 표준 헤더, validation, 서식, 안내와 통합 대시보드도 함께 보증합니다.

`setupSystem()`은 install + repair 방식입니다. 다시 실행해도 폴더, Spreadsheet, 설정 행, 헤더, trigger를 중복 생성하지 않으며 기존 업무 데이터를 초기화하지 않습니다.

## 필수 설정

`Polar Penguin` Spreadsheet의 `설정` 시트에서 확인합니다.

| 키 | 용도 | 기본/주의사항 |
|---|---|---|
| `알림이메일` | Input 처리 실패 알림 수신자 | 비어 있으면 로그만 남기고 메일은 보내지 않음 |
| `폴링주기(분)` | `processInput`, `syncAndRefresh` 실행 주기 | 기본 5분, 1/5/10/15/30분 지원 |
| `지시번호접두어` | 피킹지시번호 접두어 | 기본 `PK` |
| `예약키워드` | 예약상품 판정 | 기본 `예약`, 쉼표로 복수 지정 가능 |
| `재고경고임계치` | 대시보드 재고 경고 | 기본 3 |
| `추가투입임계(분)` | 주문 추가 투입 권고 기준 | 기본 45분 |

`파일ID`, `시트명`, `통합Input폴더ID`, `Processed폴더ID`, `Error폴더ID`, `Output폴더ID`, `Backup폴더ID`는 setup이 자동 관리합니다. 특별한 이유가 없으면 수동으로 변경하지 않습니다. `Backup`은 향후 백업/보관을 위한 예약 폴더로, 현재 파이프라인이 자동 복사하지는 않습니다.

## 일상 사용법

1. 카페24에서 주문 또는 재고 파일을 내려받습니다.
2. 파일명을 바꾸거나 주문/재고별 폴더로 나누지 말고 `Input` 직하위에 업로드합니다.
3. 설정된 폴링 주기를 기다립니다.
4. 바로 처리하려면 Spreadsheet의 `📦 Polar Penguin → 📥 Input 처리 → Input 지금 처리`를 선택하거나 Apps Script에서 `processInput()`을 실행합니다.

지원 형식은 CSV와 Google Spreadsheet입니다. Google Spreadsheet는 첫 번째 시트를 읽습니다. CSV는 UTF-8을 우선 사용하고 깨진 문자가 있으면 EUC-KR/CP949로 다시 읽습니다. Excel `.xlsx`는 지원하지 않으므로 CSV 또는 Google Spreadsheet로 변환해야 합니다.

### 파일 판별 기준

파일명은 판별에 사용하지 않습니다. 첫 행에 다음 헤더 조합이 있어야 합니다. 공백 차이는 무시합니다.

| 유형 | 필수 헤더 | 허용 별칭 |
|---|---|---|
| 주문 | `주문번호`, `품목별 주문번호`, `상품품목코드`, `수량` | `품목별주문번호`, `주문상세번호`, `품목코드`, `상품코드`, `주문수량` |
| 재고 | `품목코드`, `상품명`, `재고수량` | `상품품목코드` |

주문 수량은 0보다 큰 숫자여야 하며, 재고수량은 숫자여야 합니다. 필수값이 빈 행은 전체 파일 실패로 처리합니다.

## 자동 처리 흐름

### 재고 파일

```text
Input → 검증 → S1 재고 동기화 → S3 예약대기 재검토
      → S4 피킹지시 → S9 PDF → Processed/YYYY-MM-DD
```

여러 파일이 함께 있으면 재고 파일을 주문 파일보다 먼저 처리합니다. 예약대기 주문은 최신 재고로 재평가됩니다.

### 주문 파일

```text
Input → 검증 → S2 주문 취입 → S3 확정/예약대기/취소 판정
      → S4 피킹지시 → S9 PDF → Processed/YYYY-MM-DD
```

`품목별 주문번호`로 기존 주문 라인을 중복 제거합니다. 주문 단위 확정, 예약상품, 재고 부족, 예약재고 규칙은 기존 S3/S4 로직을 그대로 사용합니다.

### 결과 폴더와 PDF

```text
Processed/YYYY-MM-DD/  정상 처리된 원본
Error/YYYY-MM-DD/      검증 또는 업무 처리에 실패한 원본
Output/YYYY-MM-DD/     신규 피킹지시 PDF
```

PDF 파일명은 `<피킹지시번호>.pdf`입니다. Output 전체에 같은 파일명이 있으면 다시 생성하지 않습니다. 자동 PDF 생성은 피킹데이터를 읽기만 하며 피킹담당자나 출력일시를 수정하지 않습니다. 원본 파일은 삭제하지 않고 날짜별 폴더로 이동합니다.

## 피킹 결과 입력

1. `Output` PDF 또는 Spreadsheet 메뉴의 `작업지시서 출력`으로 피킹합니다.
2. `피킹(라인)`의 `확인`에 `O` 또는 `X`를 입력합니다.
3. `O`는 정상 피킹, `X`는 예외입니다. `X`일 때는 `예외사유`에 `재고없음` 또는 `불량재고`를 선택합니다.
4. 폴링 trigger가 S5 결과반영을 수행합니다. 바로 반영하려면 메뉴의 `주문/피킹 → 결과 반영`을 선택합니다.

한 주문의 품목 중 하나라도 `X`면 기존 규칙에 따라 주문 전체를 취소하고 이미 예약한 재고를 복원합니다.

## 오류 처리와 재시도

오류가 발생하면 원본을 `Error/YYYY-MM-DD`로 이동하고 `작업로그`, `입력처리로그`에 기록합니다. `알림이메일`이 설정되어 있으면 `[Polar Penguin] Input Processing Failed`라는 제목으로 오류 정보를 보냅니다.

| 오류 코드 | 의미 |
|---|---|
| `EMPTY_FILE` | 헤더 아래에 데이터가 없음 |
| `UNKNOWN_TYPE` | 주문/재고 필수 헤더 조합이 없음 |
| `MISSING_VALUE` | 데이터 행의 필수값이 비어 있음 |
| `INVALID_VALUE` | 수량 또는 재고수량이 올바른 숫자가 아님 |
| `UNSUPPORTED_FORMAT` | CSV/Google Spreadsheet가 아닌 파일 |
| `CORRUPT_FILE` | CSV 인코딩 또는 내용을 해석할 수 없음 |
| `DUPLICATE_FILE` | 이미 성공적으로 처리한 내용과 체크섬이 같음 |
| `PROCESSING_FAILED` | S1~S4/S9 처리 중 예기치 않은 오류 |

재시도하려면 오류 메시지에 따라 원본을 수정한 뒤 `Input`에 다시 업로드합니다. 실패 이력은 재시도를 막지 않습니다. 이미 `PROCESSED`로 기록된 내용을 다시 올리면 중복 반영하지 않고 `Error`로 이동합니다.

## 멱등성과 데이터 보존

- 파일 내용 SHA-256과 `입력처리로그`로 동일 파일 재처리를 막습니다.
- `품목별 주문번호`로 동일 주문 라인 중복 적재를 막습니다.
- 재고 파일은 재고 스냅샷을 기준으로 가용재고를 다시 계산합니다.
- 피킹지시번호별 PDF는 한 번만 생성합니다.
- Input 원본은 삭제하지 않고 `Processed` 또는 `Error`로 이동합니다.
- `setupSystem()` 재실행은 기존 시트나 데이터를 초기화하지 않습니다.

## 운영 점검

- `📊 대시보드`의 작업 영역은 실행 가능한 셀 버튼이 아니라 메뉴 위치를 보여주는 안내입니다. Apps Script는 일반 셀 클릭에 함수를 연결할 수 없으므로 실제 작업은 상단 커스텀 메뉴에서 실행합니다. Drawing/Image 버튼은 사용자 편집 과정에서 함수 할당이 필요해 setup이 자동 생성하지 않습니다.
- 메뉴의 `시스템 상태 확인`: 단일 Spreadsheet 연결, 필수 탭/컬럼, trigger 확인
- `📖 안내`: 현재 Input → 주문/재고 → 피킹 → O/X 결과 반영 흐름과 폴더/용어 설명
- `진단_주문폴더()`: 호환 함수명이며 통합 Input 파일의 헤더 판별 결과를 로그에 출력
- `설정 캐시 초기화`: 설정 시트 수정 후 바로 반영할 때 사용
- `setupSystem()`: 누락된 폴더, 시트, 헤더, validation, trigger를 복구

## 개발 검증

```bash
npm test
npm run lint
clasp status
```

`clasp status`에는 `src/` 안의 Apps Script 런타임 파일만 표시되어야 합니다.
