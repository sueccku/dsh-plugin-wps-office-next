// 80 场景基准的静态验收（P2-2）：不跑模型、不碰 WPS。
// 只证明「80 条都有 id/标题/应用/提示/确定性断言/coreCheck」，且没有重复/漏项。
import { BENCHMARK80 } from './scenarios-80.mjs';

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok });
  console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : ''));
}

check('80 cases', BENCHMARK80.length === 80, String(BENCHMARK80.length));
const ids = BENCHMARK80.map((s) => s.id);
const keys = BENCHMARK80.map((s) => s.key);
check('unique ids', new Set(ids).size === ids.length, ids.length + ' ids');
check('unique keys', new Set(keys).size === keys.length, keys.length + ' keys');
check('every case has title/prompt/coreCheck', BENCHMARK80.every((s) => s.title && s.prompt && s.coreCheck));
check('every case has deterministic checks', BENCHMARK80.every((s) => Array.isArray(s.checks) && s.checks.length >= 2));
check('every check uses a known verifier prefix', BENCHMARK80.every((s) => s.checks.every((c) => /^(file-exists|file-absent|docx-|xlsx-|pptx-|unchanged|transcript-)/.test(c))));
const apps = BENCHMARK80.reduce((m, s) => (m[s.app] = (m[s.app] || 0) + 1, m), {});
console.log('apps=' + JSON.stringify(apps));
check('all four app families present', ['excel', 'word', 'ppt', 'cross'].every((a) => apps[a] > 0));
check('no case is silently ready without fixture', BENCHMARK80.every((s) => s.ready || s.fixture === null));

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'BENCHMARK80 OK (' + results.length + ')' : 'BENCHMARK80 FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);
