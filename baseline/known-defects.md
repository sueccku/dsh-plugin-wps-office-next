# Known Defects (regression checklist)

Every item was verified by reading upstream source at a825336, or is documented in upstream
PR #36 (still open, unmerged). Each row is a required regression case for the MVP.

Legend: [V] verified in source, [PR] reported upstream.

Status column added 2026-09-13: **fixed** (with the docs/FIXES.md entry that closed it), **partial**,
**open**, or **wontfix (by design)**. Every status was re-checked against the current source, not
copied from the fix log.

## Transport and process

| ID | Defect | Evidence | Status |
|---|---|---|---|
| T1 | No -NoProfile / -NonInteractive on every spawn; profile stdout corrupts JSON parse | [V] wps-client.ts:87-96,138 | **fixed** — the resident host is spawned with `-NoProfile -NoLogo -NonInteractive -STA` (com-host.ts:161) and speaks JSON only over stdin/stdout |
| T2 | Fixed 30s timeout for all actions, including PDF export, pivot and image export | [V] wps-client.ts:67,106-110 | **fixed** — `timeoutFor(action)` gives per-action timeouts (com-host.ts:65,120-124), plus 330s at the MCP layer |
| T3 | Request params ride the command line; large write_range payloads hit argv limits | [V] wps-client.ts:86-92 | **fixed** — params travel as a JSON frame on the host's stdin |
| T4 | No queue or mutex; concurrent calls run concurrent COM calls | [V] wps-client.ts:83-96 | **fixed** — one serial promise chain per host (com-host.ts:93) |
| T5 | Shared mutable pptTargetName bleeds across concurrent calls | [V] wps-client.ts:72-78,338-340 | **partial** — the variable is still there (wps-client.ts:26) and is now the deliberate "lock the target deck" feature; serial dispatch removes the concurrent bleed it was reported for |
| T6 | PS_SCRIPT_PATH derived from __dirname/../../ breaks if dist is relocated | [V] wps-client.ts:38 | **fixed** — `plugin.js` publishes `WPS_OFFICE_HOST_SCRIPT` / `WPS_OFFICE_MCP_ENTRY` as absolute paths |
| T7 | wps-keepalive.ts is dead code (never imported) | [V] client/wps-keepalive.ts | **fixed** — the file no longer exists |

## COM layer

| ID | Defect | Evidence | Status |
|---|---|---|---|
| C1 | 50 bare catch blocks swallow failures, some then report success | [V] wps-com.ps1 34,631,667,735,800-804,1199-1233,1291,2700,2802 | **fixed** (FIXES 23) — 49 became `Add-WpsWarning`; 11 remain intentional and are listed there |
| C2 | Pivot refresh always records success even on failure | [V] wps-com.ps1:804 | **fixed** — a failed `RefreshTable()` now records a warning (wps-com.ps1:1235) |
| C3 | No Marshal::ReleaseComObject anywhere; leaks bounded only by process exit | [V] grep = 0 | **wontfix (by design)** — the resident host owns the COM instance and releases it at process exit; explicit release mid-host is what corrupts the cached RCW |
| C4 | getRangeData(sheet||0) maps to Sheets.Item(0) and fails when sheet is omitted | [V] excel/data.ts:71 + wps-com.ps1:577 | **fixed** (FIXES 3) — 6 sites fall back to `$excel.ActiveSheet` |
| C5 | Per-cell COM loops for range read and cleanData make operations O(n*m) round trips | [V] wps-com.ps1:580-583,664-678 | **fixed** (FIXES 4) — one `Range.Value2` read plus jagged-array conversion; 8000 cells in 37ms |
| C6 | COM getters fall back to New-Object -ComObject, silently launching a hidden WPS instance | [V] wps-com.ps1:18-86 | **partial** — the fallback survives (wps-com.ps1:76) but a candidate must pass `Test-WpsAppUsable` before use; the `[Activator]` path that produced hollow app objects is deliberately gone |
| C7 | Target re-resolved per action (TOCTOU); a window switch retargets writes | [V] wps-com.ps1:547,690 | **open** — actions still read `ActiveWorkbook`/`ActiveDocument`/`ActivePresentation`; Excel/Word write tools accept an explicit sheet/doc name, which is the mitigation we document, not a fix |
| C8 | Get-TargetPres may return null but beautifySlide dereferences it unchecked | [V] wps-com.ps1:90-96,4910-4912 | **partial** — `beautifySlide` went away with the scenario wrappers (FIXES 22); the same unchecked pattern remains at wps-com.ps1:873-874 (convertToPDF) |
| C9 | Param JSON parse failure silently becomes an empty hashtable | [V] wps-com.ps1:259 | **fixed** — the host answers `invalid request json` / `missing action` instead of continuing (wps-com-host.ps1:64,69) |
| C10 | Identical if/else branches for the sheet integer check | [V] wps-com.ps1:559,568,577 | **fixed** (FIXES 3) — the duplicated branches are gone |

## Tool layer

| ID | Defect | Evidence | Status |
|---|---|---|---|
| S1 | setActiveTarget returns success true even when validation throws | [V] ppt/presentation.ts:791-798 | **open (downgraded to honest reporting)** — the catch still returns success:true but puts the validation failure in the message body (presentation.ts:730-738) |
| S2 | Duplicate tool registration is silently skipped, hiding name collisions | [V] tool-registry.ts:68-71 | **partial** — still skipped, but now logged at warn level (tool-registry.ts:68-71); the choice is deliberate so a stale dist cannot break startup |
| S3 | Only required-parameter presence is validated; wrong parameter names are ignored silently | [V] tool-registry.ts:210-227 | **fixed** (FIXES 13/20) — required params are schema-checked and the bridge rejects unknown keys naming what the action accepts (`unknown parameter(s) ... accepted: ...`) |
| S4 | Inconsistent result shapes: builtins always success true, pro tools mix JSON and prose | [V] mcp-server.ts:175-187,205-216,321-331 | **partial** — the shape is now consistent (success + prose + warnings) but the legacy 12 builtins still always report their own shape |
| S5 | 12 builtin tools duplicate pro tools (cell value, active document, insert text) | [V] mcp-server.ts:157-678 | **open** — the 12 are still registered and show up in `wps_help`; they are not advertised, so they cost no schema bytes per request |
| S6 | Documented counts drift: tools/index.ts:23 says 235+12=247, actual is 238+12=250 | [V] tools/index.ts:23 | **fixed** — `scripts/verify.mjs` asserts the action count three ways and the budget, and `gen-skill-tools.mjs` regenerates every count in the skills |
| S7 | getAppInfo returns only the text prefix, no payload | [V] observed live | **fixed** — the handler reads structured version/build/platform/activeDocument (common/general.ts:340-366) |

## Upstream-reported (PR #36)

| ID | Defect | Evidence | Status |
|---|---|---|---|
| P1 | wps_word_set_line_spacing is registered but setLineSpacing has no COM action | [PR] returns Unknown action | **fixed** (FIXES 1) |
| P2 | wps_word_set_page_setup returns undefined (setPageSetup parameter handling bug) | [PR] | **fixed** — margins are validated as whole points 0-1584 and the effective values are reported (wps-com.ps1:2961+) |
| P3 | set_font with range=all overwrites heading fonts | [PR] | **wontfix (documented)** — unchanged COM behaviour; documented as a known pitfall in skills/wps-word/SKILL.md ("已知坑") |
| P4 | install.ps1 hangs on Read-Host when WPS lives in a versioned LOCALAPPDATA directory | [PR] | **wontfix** — this repo ships no `install.ps1`; installation is `dsh plugin add` |

## Constraints to preserve, not defects

- wps-com.ps1 carries a UTF-8 BOM and sets [Console]::OutputEncoding=UTF8; both are load-bearing.
- The 245 COM actions are the valuable asset; the MVP rewires transport and surface, not the actions.
  (The action count is now 259; the generator enforces source == generated == expected.)
