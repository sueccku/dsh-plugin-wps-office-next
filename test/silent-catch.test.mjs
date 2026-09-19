// S5: empty catch blocks are a ledger, not an accident.
//
// Every empty catch in the two hand-maintained PowerShell files must be registered here with a reason.
// A NEW empty catch fails this test until it is registered, and a REMOVED one fails until the ledger is
// tightened. The generated host/wps-actions.ps1 is not scanned: CI already asserts it is byte-identical
// to a regeneration of mcp/scripts/wps-com.ps1.
// Run: node test/silent-catch.test.mjs
import { readFileSync } from "node:fs";

/**
 * file -> { "<trimmed source line>": { count, reason } }.
 * The key is the source line containing the empty catch, normalised to single spaces, so moving the
 * code does not break the ledger but changing or adding one does.
 */
const ALLOWLIST = {
  "mcp/scripts/wps-com.ps1": {
    "if ($kind -ne 'excel') { try { $active.Visible = $true } catch { } }": { count: 1, reason: "Get-WpsApp: make a reused non-Excel instance visible; harmless if it refuses" },
    "if ($kind -ne 'excel') { try { $created.Visible = $true } catch { } }": { count: 1, reason: "Get-WpsApp: same, for the freshly created instance" },
    "} catch { }": { count: 9, reason: "function-level fallback: Get-WpsApp / Get-RangeAddressSafe / Get-ListObjectAddress / Get-WpsRangeImpact / Get-WpsAppRealVersion fall back to a default when the read fails, and the FIXES 66 ownership record (Save-WpsOwnedApps / Clear-WpsOwnedApps) plus the per-kind reclaim quit are best effort" },
    "try { $app.Quit(); $closed += $kind } catch { }": { count: 1, reason: "Close-WpsAppsStartedByUs: a failed Quit only means one instance is not released" },
    "try { $app.DisplayAlerts = $prev } catch { }": { count: 1, reason: "Restore-WpsAlerts: a failed restore degrades to silence, the action result is already in hand" },
    "try { return [int]$range.Application.WorksheetFunction.CountA($range) } catch { }": { count: 1, reason: "Get-WpsRangeNonEmpty: preferred count path; falls through on failure" },
    "try { return [int]$range.Application.CountA($range) } catch { }": { count: 1, reason: "Get-WpsRangeNonEmpty: fallback count path; falls through again" },
    "try { $impact.address = Get-RangeAddressSafe $range $null } catch { }": { count: 1, reason: "Get-WpsRangeImpact: best-effort stats degrade to null/empty" },
    "try { $impact.cells = [int]$range.Count } catch { }": { count: 1, reason: "Get-WpsRangeImpact: best-effort stats degrade to null/empty" },
    "try { $preview += [string]$doc.Comments.Item([int]$p.index).Range.Text } catch { }": { count: 1, reason: "deleteComment: capture the note text; if it cannot be read, report without a sample" },
    "try { $preview += [string]$doc.Comments.Item($i).Range.Text } catch { }": { count: 1, reason: "deleteComment: same, per-item in the loop" },
    "try { if ($cell.Comment) { $hadComment = $true; $oldText = [string]$cell.Comment.Text() } } catch { }": { count: 2, reason: "add/deleteCellComment: read the previous note text; if it cannot be read, report without a sample" },
    "try { $prevAskLinks = [bool]$excel.AskToUpdateLinks; $excel.AskToUpdateLinks = $false } catch { }": { count: 1, reason: "openWorkbook: remember and suppress AskToUpdateLinks; skip if the property is unavailable" },
    "try { if ($null -ne $prevAskLinks) { $excel.AskToUpdateLinks = $prevAskLinks } } catch { }": { count: 1, reason: "openWorkbook: restore AskToUpdateLinks; best effort" },
    "try { $prevAlerts = [bool]$excel.DisplayAlerts } catch { }": { count: 1, reason: "closeWorkbook: read/restore DisplayAlerts around the close; failure must not block the close" },
    "try { $excel.DisplayAlerts = $false } catch { }": { count: 1, reason: "closeWorkbook: read/restore DisplayAlerts around the close; failure must not block the close" },
    "try { $excel.DisplayAlerts = $prevAlerts } catch { }": { count: 1, reason: "closeWorkbook: read/restore DisplayAlerts around the close; failure must not block the close" },
    "try { $prevAlerts = [bool]$word.DisplayAlerts } catch { }": { count: 1, reason: "closeDocument: read/restore DisplayAlerts around the close; failure must not block the close" },
    "try { $word.DisplayAlerts = 0 } catch { }": { count: 1, reason: "closeDocument: read/restore DisplayAlerts around the close; failure must not block the close" },
    "try { $word.DisplayAlerts = $prevAlerts } catch { }": { count: 1, reason: "closeDocument: read/restore DisplayAlerts around the close; failure must not block the close" },
    "$ver = \"\"; try { $ver = [string]$excel.Version } catch { }": { count: 1, reason: "getAppInfo: read the running WPS version; report empty if unavailable" },
    "$bld = \"\"; try { $bld = [string]$excel.Build } catch { }": { count: 1, reason: "getAppInfo: read the running WPS build; report empty if unavailable" },
    "$ver = \"\"; try { $ver = [string]$word.Version } catch { }": { count: 1, reason: "getAppInfo: read the running WPS version; report empty if unavailable" },
    "$bld = \"\"; try { $bld = [string]$word.Build } catch { }": { count: 1, reason: "getAppInfo: read the running WPS build; report empty if unavailable" },
    "$ver = \"\"; try { $ver = [string]$ppt.Version } catch { }": { count: 1, reason: "getAppInfo: read the running WPS version; report empty if unavailable" },
    "$bld = \"\"; try { $bld = [string]$ppt.Build } catch { }": { count: 1, reason: "getAppInfo: read the running WPS build; report empty if unavailable" },
    "try { $openCount = [int]$ppt.Presentations.Count } catch { }": { count: 1, reason: "Get-TargetPres: count open presentations to warn about an ambiguous active target; skip the warning if the count cannot be read" },
    "try { $openCount = [int]$excel.Workbooks.Count } catch { }": { count: 2, reason: "Get-WorksheetByParam / Resolve-Worksheet: count open workbooks to warn about an ambiguous active target; skip the warning if the count cannot be read" },
  },
  "host/wps-com-host.ps1": {
    "} catch { }": { count: 1, reason: "Update-HostState: bookkeeping must never take an action down" },
    "try { Stop-Process -Id ([int]$state.hostPid) -Force -ErrorAction SilentlyContinue } catch { }": { count: 1, reason: "Acquire-HostMutex: stopping the stale host is best effort" },
    "try { if ($mutex.WaitOne(5000)) { return $mutex } } catch [System.Threading.AbandonedMutexException] { return $mutex } catch { }": { count: 1, reason: "Acquire-HostMutex: takeover failure ends in refusing to start, which the caller reports" },
    "try { $null = Invoke-WpsAction -Action '__warmup' -Params '{}' } catch { }": { count: 1, reason: "warmup is an optimisation; the first real call pays the JIT cost if it fails" },
    "try { $null = Close-WpsAppsStartedByUs } catch { }": { count: 1, reason: "FIXES 65 exit net: releasing our own WPS instances must never turn a clean EOF exit into a crash" },
    "try { $null = Invoke-WpsOrphanReclaim } catch { }": { count: 1, reason: "FIXES 66 reclaim is best effort; a failed reclaim must never stop the host from serving" },
  },
};

function findEmptyCatches(file) {
  const src = readFileSync(file, "utf8");
  const lines = src.split(/\r?\n/);
  const found = [];
  const re = /\bcatch\b/g;
  let m;
  while ((m = re.exec(src))) {
    let i = m.index + m[0].length;
    while (src[i] === " " || src[i] === "\t") i++;
    if (src[i] === "[") { const close = src.indexOf("]", i); if (close < 0) continue; i = close + 1; }
    while (src[i] === " " || src[i] === "\t") i++;
    if (src[i] !== "{") continue;
    let depth = 0, j = i;
    for (; j < src.length; j++) {
      if (src[j] === "{") depth++;
      else if (src[j] === "}") { depth--; if (depth === 0) { j++; break; } }
    }
    const body = src.slice(i + 1, j - 1);
    const stripped = body.replace(/^[ \t]*#.*$/gm, "").replace(/\s+/g, "");
    if (stripped === "") {
      const line = src.slice(0, m.index).split("\n").length;
      found.push(lines[line - 1].trim().replace(/\s+/g, " "));
    }
    re.lastIndex = j;
  }
  return found;
}

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

for (const file of Object.keys(ALLOWLIST)) {
  const found = findEmptyCatches(file);
  const foundCounts = {};
  for (const line of found) foundCounts[line] = (foundCounts[line] || 0) + 1;
  const allow = ALLOWLIST[file];

  const unregistered = [];
  for (const [line, n] of Object.entries(foundCounts)) {
    const a = allow[line] ? allow[line].count : 0;
    if (n > a) unregistered.push(n - a + "x " + JSON.stringify(line.slice(0, 90)));
  }
  const stale = [];
  for (const [line, entry] of Object.entries(allow)) {
    const n = foundCounts[line] || 0;
    if (entry.count > n) stale.push(JSON.stringify(line.slice(0, 90)) + " expected " + entry.count + " found " + n);
  }
  const total = found.length;
  const allowTotal = Object.values(allow).reduce((n, e) => n + e.count, 0);

  check(file + ": no unregistered empty catch", unregistered.length === 0, unregistered.length ? unregistered.join(" | ") : total + " registered");
  check(file + ": no stale allowlist entry", stale.length === 0, stale.length ? stale.join(" | ") : "all entries still present");
  check(file + ": every entry carries a reason", Object.values(allow).every((e) => typeof e.reason === "string" && e.reason.length > 10), Object.keys(allow).length + " entries");
  check(file + ": total count matches the ledger", total === allowTotal, "found=" + total + " ledger=" + allowTotal);
}

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "SILENT CATCH TESTS OK (" + results.length + ")" : "SILENT CATCH TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);
