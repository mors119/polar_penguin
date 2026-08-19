# Warehouse workflow responsibility classification

This classification was completed before the automation-first refactor and is the basis of the implemented architecture.

## Responsibilities

| Responsibility | Classification | Result |
|---|---|---|
| Input polling, file type detection, schema validation, checksum protection | AUTOMATE | `processInput` owns the complete input boundary. |
| Cafe24 inventory synchronization | AUTOMATE | S1 remains an internal service called by S6. |
| Order import and item-order deduplication | AUTOMATE | S2 remains an internal service called by S6. |
| Whole-order stock aggregation and reservation | AUTOMATE | S3 aggregates duplicate SKUs and reserves once for only the explicit new/released orders. |
| Picking header/line creation and deduplication | AUTOMATE | S4 is an explicit-target internal service. |
| Initial picking PDF creation and dashboard refresh | AUTOMATE | S6 marks the order ready only after PDF success. |
| O/X result reflection | AUTOMATE | `syncAndRefresh` processes complete order results periodically; the menu action is recovery/immediate execution only. |
| Order cancellation decision and reason | KEEP_EDITABLE | Controlled selection-based menu flow; restoration is automatic. |
| Reservation release decision | KEEP_EDITABLE | Operator selects one reservation order after reviewing availability. |
| Physical picking result | KEEP_EDITABLE | O/X and X reason are the primary line inputs. |
| Warehouse location and warehouse notes | KEEP_EDITABLE | Visually marked and never overwritten by inventory sync. |
| Picking worker identity | KEEP_EDITABLE | Entered once on the picking header and propagated to lines. |
| Order/picking IDs, states, quantities and timestamps | INTERNAL_ONLY | System-generated and visually distinguished from editable fields. |
| Inventory movement, operation and input logs | INTERNAL_ONLY | Retained for idempotency, audit and recovery. |
| Folder IDs and operational spreadsheet ID | INTERNAL_ONLY | System-managed settings; setup repairs invalid values. |
| Polling, prefix, keywords, thresholds, alert address and aliases | KEEP_EDITABLE | User-managed settings preserved across setup reruns. |
| Manual S1 inventory sync, S2 ingest, S3 confirmation, S4 picking generation | REMOVE | Removed from the operator menu; services remain for orchestration. |
| Automatic reservation release after inventory input | REMOVE | Inventory input recalculates availability only. |
| Separate status sheets and legacy manual dashboard flows | REMOVE | One source of truth plus one integrated dashboard. |

## Sheets

| Sheet | Operator role | Classification | Visibility |
|---|---|---|---|
| 📖 안내 | Current operator instructions | operator information | Visible |
| 📊 대시보드 | Actions and exceptions requiring attention | operator information | Visible |
| 상품마스터 | Warehouse location/notes plus synchronized catalog and inventory | operator data entry + system record | Visible |
| 주문(완료) | Four-state operational record; optional operator memo | operator information + limited entry | Visible |
| 피킹(라인) | O/X and exception reason | operator data entry | Visible |
| 피킹(헤더) | Order aggregation, worker entered once, output error recovery | system internal storage + limited entry | Hidden by default |
| 재고이동로그 | Inventory audit and idempotency evidence | system internal storage | Hidden by default |
| 작업로그 | Diagnostics and recovery evidence | system internal storage | Hidden by default |
| 입력처리로그 | Input checksum/idempotency and failure history | system internal storage | Hidden by default |
| 설정 | System-managed IDs and user-managed operational parameters | system internal storage + limited entry | Hidden by default |
| 예약대기, 주문반려, 주문현황, 재고현황, 피킹현황 | Redundant derived views | obsolete | Not created or maintained |

## Legacy state review

`접수`, `확정`, and `예약대기` are migrated during setup only and are never written by the current workflow. `피킹대기` and `피킹중` have no persistent representation. The only persistent order states are `처리완료`, `예약`, `출고완료`, and `취소`. Temporary progress is represented by the current function call and internal picking header values (`대기`, `완료`, `취소`, `출력오류`), not by extra order states.
