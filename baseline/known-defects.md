# Known Defects (regression checklist)

Every item was verified by reading upstream source at a825336, or is documented in upstream
PR #36 (still open, unmerged). Each row is a required regression case for the MVP.

Legend: [V] verified in source, [PR] reported upstream.

## Transport and process

| ID | Defect | Evidence |
|---|---|---|
| T1 | No -NoProfile / -NonInteractive on every spawn; profile stdout corrupts JSON parse | [V] wps-client.ts:87-96,138 |
| T2 | Fixed 30s timeout for all actions, including PDF export, pivot and image export | [V] wps-client.ts:67,106-110 |
| T3 | Request params ride the command line; large write_range payloads hit argv limits | [V] wps-client.ts:86-92 |
| T4 | No queue or mutex; concurrent calls run concurrent COM calls | [V] wps-client.ts:83-96 |
| T5 | Shared mutable pptTargetName bleeds across concurrent calls | [V] wps-client.ts:72-78,338-340 |
| T6 | PS_SCRIPT_PATH derived from __dirname/../../ breaks if dist is relocated | [V] wps-client.ts:38 |
| T7 | wps-keepalive.ts is dead code (never imported) | [V] client/wps-keepalive.ts |

## COM layer

| ID | Defect | Evidence |
|---|---|---|
| C1 | 50 bare catch blocks swallow failures, some then report success | [V] wps-com.ps1 34,631,667,735,800-804,1199-1233,1291,2700,2802 |
| C2 | Pivot refresh always records success even on failure | [V] wps-com.ps1:804 |
| C3 | No Marshal::ReleaseComObject anywhere; leaks bounded only by process exit | [V] grep = 0 |
| C4 | getRangeData(sheet||0) maps to Sheets.Item(0) and fails when sheet is omitted | [V] excel/data.ts:71 + wps-com.ps1:577 |
| C5 | Per-cell COM loops for range read and cleanData make operations O(n*m) round trips | [V] wps-com.ps1:580-583,664-678 |
| C6 | COM getters fall back to New-Object -ComObject, silently launching a hidden WPS instance | [V] wps-com.ps1:18-86 |
| C7 | Target re-resolved per action (TOCTOU); a window switch retargets writes | [V] wps-com.ps1:547,690 |
| C8 | Get-TargetPres may return null but beautifySlide dereferences it unchecked | [V] wps-com.ps1:90-96,4910-4912 |
| C9 | Param JSON parse failure silently becomes an empty hashtable | [V] wps-com.ps1:259 |
| C10 | Identical if/else branches for the sheet integer check | [V] wps-com.ps1:559,568,577 |

## Tool layer

| ID | Defect | Evidence |
|---|---|---|
| S1 | setActiveTarget returns success true even when validation throws | [V] ppt/presentation.ts:791-798 |
| S2 | Duplicate tool registration is silently skipped, hiding name collisions | [V] tool-registry.ts:68-71 |
| S3 | Only required-parameter presence is validated; wrong parameter names are ignored silently | [V] tool-registry.ts:210-227 |
| S4 | Inconsistent result shapes: builtins always success true, pro tools mix JSON and prose | [V] mcp-server.ts:175-187,205-216,321-331 |
| S5 | 12 builtin tools duplicate pro tools (cell value, active document, insert text) | [V] mcp-server.ts:157-678 |
| S6 | Documented counts drift: tools/index.ts:23 says 235+12=247, actual is 238+12=250 | [V] tools/index.ts:23 |
| S7 | getAppInfo returns only the text prefix, no payload | [V] observed live |

## Upstream-reported (PR #36)

| ID | Defect | Evidence |
|---|---|---|
| P1 | wps_word_set_line_spacing is registered but setLineSpacing has no COM action | [PR] returns Unknown action |
| P2 | wps_word_set_page_setup returns undefined (setPageSetup parameter handling bug) | [PR] |
| P3 | set_font with range=all overwrites heading fonts | [PR] |
| P4 | install.ps1 hangs on Read-Host when WPS lives in a versioned LOCALAPPDATA directory | [PR] |

## Constraints to preserve, not defects

- wps-com.ps1 carries a UTF-8 BOM and sets [Console]::OutputEncoding=UTF8; both are load-bearing.
- The 245 COM actions are the valuable asset; the MVP rewires transport and surface, not the actions.
