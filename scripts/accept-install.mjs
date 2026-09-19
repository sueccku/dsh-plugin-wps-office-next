// 安装验收：在一个全新的、一次性的 profile 里，把本仓库当成正式包完整装一遍再拆掉。
//
// 为什么需要它：README 让用户把一段话交给 AI 去装，而这条路径此前只在本机既有的 web profile 上
// 手工验证过。一个干净的 profile 能挡住「只在旧 profile 的残留产物里才成立」的假通过。
//
// 它做四件事：
//   1. dsh plugin add 本地仓库（不走网络）
//   2. dump-config 里必须出现 wps-office-next-plugin 与 mcp-wps-office-next 两个 id
//   3. 装出来的副本自带预构建产物，且能从这个副本里跑通 scripts/doctor.mjs
//   4. 卸载干净（默认连一次性 profile 目录一起删掉，--keep 可以留下来看）
//
// Usage: node scripts/accept-install.mjs [--profile <name>] [--keep]
import { spawnSync } from 'node:child_process';
import { existsSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const DSH_HOME = process.env.DSH_HOME || join(homedir(), '.dsh');
const argv = process.argv.slice(2);
const keep = argv.includes('--keep');
const profileArg = argv.indexOf('--profile');
const profile = profileArg >= 0 && argv[profileArg + 1] ? argv[profileArg + 1] : 'wps-acceptance';

const checks = [];
function check(name, ok, detail) { checks.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
function run(command, args, options = {}) {
  // shell:true 是为了让 dsh 的 .cmd/.ps1 垫片能被解析；带空格的 node 路径必须走 shell:false。
  const result = spawnSync(command, args, { encoding: 'utf8', shell: options.shell !== false, timeout: 300000, ...options });
  return { status: result.status, stdout: String(result.stdout || ''), stderr: String(result.stderr || ''), error: result.error };
}

const profileDir = join(DSH_HOME, 'profiles', profile);
const existedBefore = existsSync(profileDir);
console.log('profile: ' + profile + '  (' + profileDir + ')' + (existedBefore ? '  [已存在，收尾时不会自动删除]' : '  [全新]'));

// 1) 安装
const add = run('dsh', ['plugin', '--profile', profile, 'add', ROOT]);
check('dsh plugin add succeeds', add.status === 0, (add.stderr || add.stdout).trim().split(/\r?\n/).slice(-2).join(' | ').slice(0, 160));

// 2) 接线
const dump = run('dsh', ['--profile', profile, '--dump-config']);
const dumpText = dump.stdout + dump.stderr;
check('dump-config reports the plugin entry', dumpText.includes('wps-office-next-plugin'), dump.status === 0 ? '' : 'exit=' + dump.status);
check('dump-config reports the MCP client', dumpText.includes('mcp-wps-office-next'), dump.status === 0 ? '' : 'exit=' + dump.status);

// 3) 装出来的副本自身要完整、能自检
const installed = join(profileDir, 'node_modules', 'dsh-plugin-wps-office-next');
check('installed copy exists', existsSync(installed), installed);
if (existsSync(installed)) {
  for (const rel of ['plugin.js', 'cordis.patch.yml', 'mcp/dist/index.js', 'host/wps-com-host.ps1', 'host/wps-actions.ps1', 'skills/wps-excel/SKILL.md']) {
    check('installed copy carries ' + rel, existsSync(join(installed, rel)), '');
  }
  const doctor = run(process.execPath, [join('scripts', 'doctor.mjs')], { cwd: installed, shell: false });
  const doctorText = doctor.stdout + doctor.stderr;
  check('doctor runs from the installed copy', doctor.status === 0, doctorText.trim().split(/\r?\n/).slice(-1)[0].slice(0, 120));
  check('doctor reports DOCTOR OK', /DOCTOR OK/.test(doctorText), '');
}

// 4) 卸载与清理
const remove = run('dsh', ['plugin', '--profile', profile, 'remove', 'dsh-plugin-wps-office-next']);
check('dsh plugin remove succeeds', remove.status === 0, (remove.stderr || remove.stdout).trim().split(/\r?\n/).slice(-1)[0].slice(0, 140));
if (!keep && !existedBefore) {
  rmSync(profileDir, { recursive: true, force: true });
  check('temporary profile removed', !existsSync(profileDir), profileDir);
} else {
  console.log('KEPT ' + profileDir + (keep ? ' (--keep)' : ' (profile existed before this run)'));
}

const failed = checks.filter((c) => !c.ok).length;
console.log(failed === 0 ? 'ACCEPT-INSTALL OK (' + checks.length + ' checks)' : 'ACCEPT-INSTALL FAILED (' + failed + '/' + checks.length + ')');
process.exit(failed === 0 ? 0 : 1);
