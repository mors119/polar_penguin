# Polar Penguin

Google Apps Script 기반 주문·재고·피킹 운영 시스템입니다.

## 설치

Apps Script의 Script Properties에 `ROOT_FOLDER_ID` 또는 `ROOT_FOLDER_URL`을 등록하거나 다음 함수를 실행합니다.

```javascript
setRootFolder('https://drive.google.com/drive/folders/...');
setupSystem();
```

`setupSystem()`은 ROOT 폴더 아래에 다음 구조를 생성하거나 기존 리소스를 재사용합니다.

```text
01 Console/Polar Penguin Console
02 Master/상품마스터
03 Orders/주문완료
04 Picking/피킹헤더, 피킹라인
Input/
Processed/YYYY-MM-DD/
Error/YYYY-MM-DD/
Output/YYYY-MM-DD/
Backup/
```

Console의 설정 시트에는 생성된 Spreadsheet와 폴더 ID가 자동 등록됩니다. 오류 메일을 받으려면 `알림이메일`을 설정합니다. 설치 함수를 다시 실행하면 누락된 구성만 복구하며 기존 설정과 업무 데이터는 유지합니다.

## 통합 Input

사용자는 주문/재고 구분 없이 `Input` 폴더에 CSV 또는 Google Spreadsheet를 올리면 됩니다. `processInput()` 트리거가 헤더로 유형을 판별하고, 기존 S1~S4를 실행한 뒤 신규 피킹지시를 PDF로 생성합니다. 정상 원본은 `Processed`, 실패 원본은 `Error`의 날짜별 폴더로 이동합니다.

## 개발 검증

```bash
npm test
npm run lint
clasp status
```
