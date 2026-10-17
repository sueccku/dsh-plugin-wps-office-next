// Skill 质量门禁（不需要 WPS）。
//
// 为什么要有它：Skill 以前靠"禁止若干词 / 限定行数"来判定，而且工具清单被生成器重新注入 SKILL.md，
// 于是"读一次技能"等于读一份工具目录。这条门禁改看**可验证的事实**：
//   1. 四个注册名与元数据齐全；
//   2. 生成块里不再有工具名（防止下一次生成把长清单塞回来）；
//   3. 正文里引用的每个相对文件真实存在；
//   4. 正文里出现的每个 wps_* 工具名都在真实注册目录里；
//   5. 单行调用示例的参数名都在该工具的 runtime schema 里；
//   6. 已被实现推翻的误导说法不再出现；
//   7. 生成器 --check 是只读的，且跑两次没有额外变化。
//
// Run: node test/skill-quality.test.mjs
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const results = [];
function check(name, ok, detail) { results.push({ name, ok: !!ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
const read = (p) => readFileSync(p, 'utf8');

const definitions = JSON.parse(read('spec/tool-definitions.json'));
const schemaByTool = new Map(definitions.map((t) => [t.name, t.inputSchema || {}]));
const known = new Set(schemaByTool.keys());

const SKILLS = [
  {
    name: 'wps-office-next',
    dir: 'skills/wps-office-next',
    // 共享协议必须落在这个文件里：应用技能被单独加载时它是唯一的全文来源。
    mustMention: ['documentId', 'wps_call', 'shortfalls', 'wps_help', 'reference.md'],
  },
  {
    name: 'wps-excel',
    dir: 'skills/wps-excel',
    mustMention: ['reference.md', 'wps-office-next', '公式', '保存'],
  },
  {
    name: 'wps-word',
    dir: 'skills/wps-word',
    mustMention: ['reference.md', 'wps-office-next', 'apply_style', '保存'],
  },
  {
    name: 'wps-ppt',
    dir: 'skills/wps-ppt',
    mustMention: ['reference.md', 'wps-office-next', 'imagePath', '磅'],
  },
];

// ---- 1) 元数据 + 应用技能里的共享协议摘要 ------------------------------------------------
for (const skill of SKILLS) {
  const file = join(skill.dir, 'SKILL.md');
  const raw = existsSync(file) ? read(file) : '';
  check(skill.name + ': SKILL.md exists', raw.length > 0, file);
  const fm = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/.exec(raw);
  check(skill.name + ': frontmatter parses', !!fm, fm ? 'ok' : 'missing frontmatter');
  const meta = {};
  if (fm) for (const line of fm[1].split(/\r?\n/)) {
    const pair = /^([A-Za-z][A-Za-z0-9_-]*):\s*(.*)$/.exec(line);
    if (pair) meta[pair[1]] = pair[2].trim();
  }
  check(skill.name + ': registered name unchanged', meta.name === skill.name, String(meta.name));
  check(skill.name + ': description present', (meta.description || '').length > 20, (meta.description || '').slice(0, 60));
  check(skill.name + ': whenToUse present', (meta.whenToUse || '').length > 10, (meta.whenToUse || '').slice(0, 60));
  for (const token of skill.mustMention) {
    check(skill.name + ': mentions ' + token, raw.includes(token), '');
  }
}

// ---- 2) 生成块里不许再有工具名 -----------------------------------------------------------
const MARK_START = '<!-- GENERATED:advertised:start -->';
const MARK_END = '<!-- GENERATED:advertised:end -->';
for (const skill of SKILLS) {
  const raw = read(join(skill.dir, 'SKILL.md'));
  const start = raw.indexOf(MARK_START);
  const end = raw.indexOf(MARK_END);
  check(skill.name + ': has exactly one generated block', start >= 0 && end > start && raw.indexOf(MARK_START, start + 1) < 0, '');
  const block = start >= 0 && end > start ? raw.slice(start + MARK_START.length, end) : '';
  // 四个门面工具名是稳定且单独出现的（不构成"目录"），其余 wps_* 一律不许出现在生成块里。
  const FACADE = new Set(['wps_status', 'wps_help', 'wps_call', 'wps_batch']);
  const names = (block.match(/wps_[a-z0-9_]+/g) || []).filter((n) => !FACADE.has(n));
  check(skill.name + ': generated block carries no tool list', names.length === 0, names.slice(0, 5).join(', ') || 'clean');
  check(skill.name + ': generated block is short', block.trim().split(/\r?\n/).length <= 6, block.trim().split(/\r?\n/).length + ' lines');
  // 手写正文里也不该再出现整段工具清单（一行 6 个以上工具名 = 又抄了一份目录）。
  const body = raw.replace(block, '');
  const longListLine = body.split(/\r?\n/).find((line) => (line.match(/wps_[a-z0-9_]+/g) || []).length >= 6);
  check(skill.name + ': no hand-written tool catalogue line', !longListLine, longListLine ? longListLine.slice(0, 80) : '');
  check(skill.name + ': body is a task guide, not a dump', body.split(/\r?\n/).length <= 200, body.split(/\r?\n/).length + ' lines');
}

// ---- 3) 相对引用必须真实存在 -------------------------------------------------------------
const allSkillFiles = [];
for (const skill of SKILLS) {
  const walk = (dir) => readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = join(dir, e.name);
    return e.isDirectory() ? walk(p) : [p];
  });
  allSkillFiles.push(...walk(skill.dir));
}
const mdFiles = allSkillFiles.filter((f) => f.endsWith('.md'));
check('skill markdown files discovered', mdFiles.length >= 12, mdFiles.length + ' files');

for (const skill of SKILLS) {
  const raw = read(join(skill.dir, 'SKILL.md'));
  const referenced = [...new Set([...raw.matchAll(/(?:[\w-]+\/)*[\w-]+\.md/g)].map((m) => m[0]))];
  for (const rel of referenced) {
    // 只校验指向本 Skill 目录的相对路径（reference.md / references/*.md / recipes/*.md）。
    if (rel !== 'reference.md' && !rel.startsWith('references/') && !rel.startsWith('recipes/')) continue;
    const target = join(skill.dir, rel);
    check(skill.name + ': reference exists -> ' + rel, existsSync(target) && statSync(target).size > 200, target);
  }
  for (const sub of ['references', 'recipes']) {
    const dir = join(skill.dir, sub);
    const files = existsSync(dir) ? readdirSync(dir).filter((f) => f.endsWith('.md')) : [];
    check(skill.name + ': ' + sub + '/ has content', files.length > 0 && files.every((f) => statSync(join(dir, f)).size > 300), files.join(', ') || 'empty');
  }
}

// ---- 4) 正文提到的每个工具名都必须在注册目录里 -------------------------------------------
for (const file of mdFiles) {
  if (file.includes('reference.md')) continue;             // 自动生成，另有门禁
  const raw = read(file);
  const refs = [...new Set([...raw.matchAll(/wps_[a-z0-9_]+/g)].map((m) => m[0]))];
  const missing = refs.filter((r) => !known.has(r));
  check('every wps_* name in ' + file + ' is registered', missing.length === 0, missing.join(', ') || refs.length + ' names');
}
// 自动生成的 reference.md 同样只允许出现真实名字。
for (const file of mdFiles) {
  if (!file.includes('reference.md')) continue;
  const refs = [...new Set([...read(file).matchAll(/wps_[a-z0-9_]+/g)].map((m) => m[0]))];
  const missing = refs.filter((r) => !known.has(r));
  check('generated ' + file + ' lists only registered tools', missing.length === 0, missing.join(', ') || refs.length + ' names');
}

// ---- 5) 单行调用示例的参数名必须在 runtime schema 里 --------------------------------------
/** 从一段 "tool { ... }" 文本里取出顶层键；跳过字符串内的括号与冒号。 */
function topLevelKeys(text) {
  const keys = [];
  let depth = 0;
  let inString = null;
  let i = 0;
  let seenBrace = false;
  while (i < text.length) {
    const c = text[i];
    if (inString) {
      if (c === '\\') { i += 2; continue; }
      if (c === inString) inString = null;
      i++;
      continue;
    }
    if (c === '"' || c === "'") { inString = c; i++; continue; }
    if (c === '{' || c === '[') { depth++; seenBrace = true; i++; continue; }
    if (c === '}' || c === ']') { depth--; i++; continue; }
    if (seenBrace && depth === 1) {
      const m = /^([A-Za-z_][A-Za-z0-9_]*)\s*:/.exec(text.slice(i));
      if (m) { keys.push(m[1]); i += m[0].length; continue; }
    }
    i++;
  }
  return keys;
}

/** 花括号净深度；跳过字符串里的括号，避免 "通知书_{部门}" 这类文本干扰。 */
function braceBalance(text) {
  let depth = 0;
  let inString = null;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inString) {
      if (c === '\\') { i++; continue; }
      if (c === inString) inString = null;
      continue;
    }
    if (c === '"' || c === "'") { inString = c; continue; }
    if (c === '{') depth++;
    else if (c === '}') depth--;
  }
  return depth;
}

/** 收集形如 "wps_tool { ... }" 的示例（允许跨行，直到括号闭合）。 */
function examples(raw) {
  const out = [];
  const lines = raw.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const m = /^\s*(?:>\s*)?(wps_[a-z0-9_]+)\s*\{/.exec(lines[i]);
    if (!m) continue;
    let text = lines[i].slice(m[0].length - 1);
    let j = i;
    for (let guard = 0; guard < 40 && braceBalance(text) > 0; guard++) {
      j++;
      if (j >= lines.length) break;
      text += '\n' + lines[j];
    }
    out.push({ tool: m[1], text, line: i + 1 });
    i = j;
  }
  return out;
}

let exampleCount = 0;
let badExample = 0;
for (const file of mdFiles) {
  if (file.includes('reference.md')) continue;
  for (const ex of examples(read(file))) {
    exampleCount++;
    if (!known.has(ex.tool)) { badExample++; check('example tool registered (' + file + ':' + ex.line + ')', false, ex.tool); continue; }
    const props = (schemaByTool.get(ex.tool) || {}).properties || {};
    const declared = new Set(Object.keys(props));
    const used = [...new Set(topLevelKeys(ex.text))];
    const unknownKeys = used.filter((k) => !declared.has(k));
    if (unknownKeys.length) {
      badExample++;
      check('example params match schema (' + file + ':' + ex.line + ' ' + ex.tool + ')', false, unknownKeys.join(', ') + ' | accepted: ' + [...declared].join(', '));
    }
  }
}
check('recipe/example parameter check ran', exampleCount >= 10 && badExample === 0, exampleCount + ' examples, ' + badExample + ' with unknown params');

// ---- 6) 已被实现推翻的说法不许再出现 -----------------------------------------------------
const BANNED = [
  '可选参数写错名不会被发现',            // 顶层未知参数现在会被拒绝
  '可选参数拼错是静默忽略',
  '传了不会被发现',
  '保留可回退版本',                      // save 会覆盖磁盘原件，不是备份
  '没连上就不要继续',                    // open/create 会启动 WPS
  '也可以按全名直接调用',                // 模型只能调用 tools/list 里的名字
  '按错误里列出的 accepted 名字重试',    // accepted 是桥接受的键，改调用参数修不了
  '演示是三个应用里工具最多的一个',
  '插入图片用 path',
  '没有 x、y 参数',
  '别把没保存的文档留在用户机器上',
];
for (const file of mdFiles) {
  const raw = read(file);
  const hits = BANNED.filter((phrase) => raw.includes(phrase));
  check('no contradicting claim in ' + file, hits.length === 0, hits.join(' / ') || 'clean');
}

// ---- 7) 生成器 --check 只读且幂等 ---------------------------------------------------------
const snapshot = new Map(allSkillFiles.filter((f) => f.endsWith('.md')).map((f) => [f, readFileSync(f)]));
const runCheck = () => spawnSync(process.execPath, ['scripts/gen-skill-tools.mjs', '--check'], { encoding: 'utf8', windowsHide: true });
const first = runCheck();
check('gen-skill-tools --check passes', first.status === 0, (first.stderr || first.stdout || '').trim().split('\n').slice(-2).join(' | '));
const second = runCheck();
check('gen-skill-tools --check is stable across two runs', second.status === 0, 'exit=' + second.status);
const changed = [...snapshot].filter(([f, before]) => !readFileSync(f).equals(before)).map(([f]) => f);
check('--check does not touch the workspace', changed.length === 0, changed.join(', ') || 'no writes');

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'SKILL QUALITY TESTS OK (' + results.length + ')' : 'SKILL QUALITY TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);
