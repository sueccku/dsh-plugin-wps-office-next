// FIXES 70：DisplayAlerts 守卫的门禁（scripts/lint-alerts.mjs）自己也要有断言——只有把「门禁会红」
// 也证明一遍，它才不是摆设。全部是纯文本分析，不需要 WPS。
// Run: node test/alerts-gate.test.mjs
import { readFileSync } from 'node:fs';
import { ALERTS_ACTION_EXCEPTIONS, collectAlertsViolations } from '../scripts/lint-alerts.mjs';

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }

const HELPERS = [
  '# 注释里提 DisplayAlerts 不算违规',
  'function Set-WpsAlertsSuppressed($app, [string]$kind) {',
  '    $prev = $null',
  '    try { $prev = $app.DisplayAlerts } catch { return $null }',
  '    try { $app.DisplayAlerts = $false } catch { return $null }',
  '    return $prev',
  '}',
  'function Restore-WpsAlerts($app, $prev) {',
  '    if ($null -eq $app -or $null -eq $prev) { return }',
  '    try { $app.DisplayAlerts = $prev } catch { }',
  '}',
  '',
].join('\n');
const action = (name, body) => '    "' + name + '" {\n' + body.map((l) => '        ' + l).join('\n') + '\n    }\n';
const scan = (src) => collectAlertsViolations(src, 'synthetic.ps1');
const rules = (v) => v.map((x) => x.rule);

// 1-2) 有弹框调用、没有守卫 —— 必须报
const unguarded = HELPERS + action('saveAs', ['$wb.SaveAs($path)']);
const v1 = scan(unguarded);
check('an unguarded SaveAs is reported', v1.length === 1 && v1[0].rule === 'alerts-missing', JSON.stringify(rules(v1)));
check('the report names the action', /saveAs/.test(v1[0] ? v1[0].detail : ''), v1[0] ? v1[0].detail.slice(0, 80) : 'no violation');

// 3) 守卫齐全 —— 必须干净
const guarded = HELPERS + action('saveAs', [
  "$prevAlerts = Set-WpsAlertsSuppressed $excel 'excel'",
  'try { $wb.SaveAs($path) } finally { Restore-WpsAlerts $excel $prevAlerts }',
]);
check('a guarded action is clean', scan(guarded).length === 0, JSON.stringify(rules(scan(guarded))));

// 4) 关了没还原 —— 必须报
const leaked = HELPERS + action('closeWorkbook', [
  "$prevAlerts = Set-WpsAlertsSuppressed $excel 'excel'",
  '$wb.Close($true)',
]);
check('a guard without its restore is reported', scan(leaked).some((v) => v.rule === 'alerts-unbalanced'), JSON.stringify(rules(scan(leaked))));

// 5) 绕过助手直接写 DisplayAlerts —— 必须报
const bypass = HELPERS + action('deleteSheet', ['$excel.DisplayAlerts = $false', '$sheet.Delete()']);
check('a direct DisplayAlerts write is reported', scan(bypass).some((v) => v.rule === 'alerts-bypassed'), JSON.stringify(rules(scan(bypass))));

// 6) 注释里的 DisplayAlerts 不算
const commented = HELPERS + action('noop', ['# 这里曾经写死 $excel.DisplayAlerts = $true']);
check('a comment mentioning DisplayAlerts is fine', scan(commented).length === 0, JSON.stringify(rules(scan(commented))));

// 7) .Open( 明确不在这条门禁里（密码框不吃 DisplayAlerts，由文件头预检挡）
const opener = HELPERS + action('legacyOpen', ['$word.Documents.Open($path)']);
check('.Open( is not treated as an alerts call', scan(opener).length === 0, JSON.stringify(rules(scan(opener))));

// 8) 过期的例外必须报
ALERTS_ACTION_EXCEPTIONS.push({ action: 'definitely_not_an_action', reason: 'test' });
const stale = scan(HELPERS + action('noop', ['$wb.SaveAs($p)']));
ALERTS_ACTION_EXCEPTIONS.pop();
check('a stale exception is reported', stale.some((v) => v.rule === 'alerts-stale-exception'), JSON.stringify(rules(stale)));

// 9-10) 真实桥：现在的源码必须干净；把任意一个守卫拿掉必须变红（证明它真的在读那份文件）
const real = readFileSync('mcp/scripts/wps-com.ps1', 'utf8');
const realViolations = collectAlertsViolations(real, 'mcp/scripts/wps-com.ps1');
check('the real bridge passes the gate', realViolations.length === 0, JSON.stringify(realViolations.slice(0, 3)));
const broken = real.replace("$prevAlerts = Set-WpsAlertsSuppressed $excel 'excel'", '$prevAlerts = $null');
check('removing one guard turns the real bridge red', broken !== real && collectAlertsViolations(broken, 'broken').length > 0, 'changed=' + (broken !== real));

const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'ALERTS GATE TESTS FAILED' : 'ALERTS GATE TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);
