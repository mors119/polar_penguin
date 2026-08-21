# Warehouse workflow responsibility classification

## Responsibilities

| Responsibility | Classification | Result |
|---|---|---|
| Input polling, validation, checksum protection | AUTOMATE | `processInput` owns the complete input boundary and moves an order source to Success only after output finalization succeeds. |
| Whole-order stock aggregation and commitment | AUTOMATE | S3 changes available stock exactly once; `재고관리=F` may naturally produce a negative net position. |
| Picking header/line creation | AUTOMATE | S4 creates audit data once and reuses an existing instruction on retry. |
| First successful PDF/manual printable output | AUTOMATE | `finalizePickingAfterOutput_` completes lines, headers, and orders without another inventory mutation. |
| Reprint | AUTOMATE | Document output only; completed state prevents inventory, status, and log mutation. |
| Order cancellation decision and reason | KEEP_EDITABLE | Selection-based cancellation is the primary correction flow; restoration is automatic and idempotent. |
| Reservation release decision | KEEP_EDITABLE | Operator selects a SKU derived from waiting order rows; the server calculates whole-order FIFO eligibility. |
| Physical picking result confirmation | REMOVE | No O/X entry or S5 confirmation is required. Successful output is the fulfillment-state point. |
| Picking line confirmation value | INTERNAL_ONLY | O is automatically stored as historical/audit data after successful output. |
| Inventory movement, operation and input logs | INTERNAL_ONLY | Retained for audit, diagnosis, and duplicate-movement protection. |
| Manual S1–S5 workflow and result polling | REMOVE | Internal services remain where needed, but no manual result menu or polling trigger is installed. |

## Sheets

| Sheet | Operator role | Classification | Visibility |
|---|---|---|---|
| 📖 안내 | Current output-commit and cancellation instructions | operator information | Visible |
| 📊 대시보드 | Reservation, shipment, cancellation, output error, and stock KPIs | operator information | Visible |
| 상품마스터 | Warehouse locations plus synchronized inventory | mixed system/operator record | Visible |
| 주문(완료) | `예약`/`출고완료`/`취소` source of truth | system record with optional memo | Visible |
| 피킹(라인) | Automatically completed picking history | system audit record | Visible |
| 피킹(헤더) | Output state and error recovery | system internal storage | Hidden by default |
| 재고이동로그 / 작업로그 / 입력처리로그 | Audit and recovery evidence | system internal storage | Hidden by default |

## State review

The persistent order states are `예약`, `출고완료`, and `취소`. Setup migrates the obsolete `처리완료` gap state back to `예약`; a successful retry then completes it through the shared output finalizer. Picking headers use `대기`, `완료`, `취소`, and `출력오류`. Header and line progression is system-managed, not an operator input workflow.
