// 门禁自检：COM 边界的「几何成员必须走 Set-ComValue」规则（S3 场景新增）自己也要有断言 —
// 只有把「门禁会红」也证明一遍，它才不是摆设。全部是纯文本分析，不需要 WPS。
// Run: node test/com-boundary-gate.test.mjs
import { readFileSync } from 'node:fs';
import { collectBoundaryViolations } from '../scripts/lint-com-boundary.mjs';

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }

const BRIDGE = 'mcp/scripts/wps-com.ps1';
const scan = (src) => collectBoundaryViolations(src, 'synthetic.ps1');
// synthetic 片段里没有账本条目，bare-continue-stale 一定报 —— 这里只关心几何规则。
const geom = (src) => scan(src).filter((v) => v.rule === 'geometry-write');
const rules = (v) => v.map((x) => x.rule);

// 1) S3 真机撞上的原写法：必须报
const prefix = geom('        if ($null -ne $p.left) { $shape.Left = $p.left }');
check('a direct geometry write is reported', prefix.length === 1 && prefix[0].rule === 'geometry-write', JSON.stringify(rules(prefix)));
check('the report names the member', /\.Left/.test(prefix[0] ? prefix[0].detail : ''), prefix[0] ? prefix[0].detail.slice(0, 90) : 'no violation');

// 2) 同族成员一个都不能漏
for (const member of ['Top', 'Width', 'Height', 'Rotation', 'ScaleWidth', 'ScaleHeight']) {
  const v = geom('        $shape.' + member + ' = $p.' + member.toLowerCase());
  check('direct ' + member + ' write is reported', v.length === 1, JSON.stringify(rules(v)));
}

// 3) 修复后的写法必须干净
const fixed = geom("        if ($null -ne $p.left) { Set-ComValue $shape 'Left' $p.left }");
check('Set-ComValue form is clean', fixed.length === 0, JSON.stringify(rules(fixed)));

// 4) 显式数值强转也是安全的（值是 Double，COM 不挑）——不能误报
const cast = geom('        if ($null -ne $p.width) { $ps.SlideWidth = [double]$p.width * 0.75 }');
check('an explicit [double] cast is not reported', cast.length === 0, JSON.stringify(rules(cast)));

// 5) 注释里提到不算违规
const commented = geom('# if ($null -ne $p.left) { $shape.Left = $p.left }');
check('a comment is not reported', commented.length === 0, JSON.stringify(rules(commented)));

// 6) 比对赋值不能被当成写入
const compare = geom('        if ($shape.Left == $p.left) { $same = $true }');
check('a comparison is not reported', compare.length === 0, JSON.stringify(rules(compare)));

// 7-8) 真实桥：现在的源码必须干净；把任意一处改回原写法必须变红（证明它真的在读那份文件）
const real = readFileSync(BRIDGE, 'utf8');
const realViolations = collectBoundaryViolations(real, BRIDGE);
check('the real bridge passes the geometry gate', realViolations.length === 0, JSON.stringify(realViolations.slice(0, 3)));
const broken = real.replace("Set-ComValue $shape 'Left' $p.left", '$shape.Left = $p.left');
check('reverting one geometry write turns the real bridge red',
  broken !== real && collectBoundaryViolations(broken, 'broken.ps1').some((v) => v.rule === 'geometry-write'),
  'changed=' + (broken !== real));

// 9-10) 原有规则不能被这次改动削弱
const range = scan('function Get-Thing($ws) {\n    $range = $ws.Range("A1")\n    return $range\n}');
check('the existing range-return rule still fires', range.some((v) => v.rule === 'range-return'), JSON.stringify(rules(range)));
check('the real bridge still has no range/bare-continue violations',
  realViolations.filter((v) => v.rule === 'range-return' || v.rule === 'bare-continue').length === 0, '');

const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'COM BOUNDARY GATE TESTS FAILED' : 'COM BOUNDARY GATE TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);
