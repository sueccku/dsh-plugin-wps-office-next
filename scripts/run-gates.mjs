// 静态门禁的**唯一入口**。命令清单只在本文件维护：AGENTS.md / README.md / docs/release-checklist.md /
// .github/workflows/ci.yml 都只引用分组，不再各抄一份命令。
//
// 背景（2026-10-18）：这三份文档 + CI 各有一份门禁清单，已经漂过两次（README 的 verify 检查数、
// 静态测试文件数）。单点维护后，加/删一道门禁只改这里。
//
// 用法：
//   node scripts/run-gates.mjs [group]
//   npm run gates -- <group>
//   group = docs | verify | code | bridge | package | release（默认）| ci（= release）
//
// 分组语义：
//   docs     只改文档：lint（含 Markdown 链接检查）+ 文档数字对账
//   verify   只读核对（**保证不写工作区**）：lint / verify --static / 参数契约只读 / spec 复现 /
//            技能参考表只读 / 静态测试清单 / 数字对账
//   code     TS / 工具 / spec：tsc + dist 对账、生成器 + 对账、lint、verify、参数契约（写入 + 对账）、
//            spec 复现、test:static、数字对账
//   bridge   code + 重建 host/wps-actions.ps1 + 对账
//   package  code + 打包冒烟 verify:package
//   release  全部（CI 跑的就是这一组；含生成器写入与对账）
import { spawnSync } from "node:child_process";
import { join } from "node:path";

const ROOT = process.cwd();
const requested = (process.argv[2] || "release").replace(/^--/, "");
const GROUP = requested === "ci" ? "release" : requested;

const node = (script, ...args) => ({ cmd: process.execPath, args: [script, ...args] });
const gitDiff = (...paths) => ({ cmd: "git", args: ["diff", "--exit-code", "--", ...paths] });
const POWERSHELL = join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const TSC = join(ROOT, "mcp", "node_modules", "typescript", "bin", "tsc");

// 组名展开：后面的分组都包含 code 的所有步骤。
const CODE = ["code", "bridge", "package", "release"];
const ALL = ["docs", "verify", ...CODE];

// 顺序与 CI 一致：先编译，再重建生成物并对账，最后跑与生成物无关的静态测试和数字对账。
const STEPS = [
  { id: "tsc", title: "tsc（mcp/dist 与源码一致）", ...node(TSC), cwd: join(ROOT, "mcp"), groups: CODE },
  { id: "dist-diff", title: "mcp/dist 对账", ...gitDiff("mcp/dist"), groups: CODE },

  { id: "build-host", title: "重建 host/wps-actions.ps1", cmd: POWERSHELL, args: ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "scripts/build-host-actions.ps1"], groups: ["bridge", "release"] },
  { id: "host-diff", title: "host/wps-actions.ps1 对账", ...gitDiff("host/wps-actions.ps1"), groups: ["bridge", "release"] },

  { id: "skill-tools-check", title: "技能参考表只读核对", ...node("scripts/gen-skill-tools.mjs", "--check"), groups: ["verify", "code", "bridge", "package"] },
  { id: "skill-tools", title: "重建技能参考表", ...node("scripts/gen-skill-tools.mjs"), groups: ["release"] },
  { id: "skills-diff", title: "skills 对账", ...gitDiff("skills"), groups: ["release"] },

  { id: "tool-surface", title: "重建 spec/*.json", ...node("scripts/gen-tool-surface.mjs"), groups: CODE },
  { id: "spec-diff", title: "spec 对账", ...gitDiff("spec"), groups: CODE },
  { id: "tool-coverage", title: "重建工具覆盖矩阵", ...node("scripts/gen-tool-coverage.mjs"), groups: CODE },
  { id: "coverage-diff", title: "docs/tool-coverage.md 对账", ...gitDiff("docs/tool-coverage.md"), groups: CODE },

  { id: "lint", title: "lint（含 Markdown 链接检查）", ...node("scripts/lint.mjs"), groups: ALL },
  { id: "fixes-index-check", title: "FIXES 索引只读核对", ...node("scripts/gen-fixes-index.mjs", "--check"), groups: ALL },
  { id: "verify-static", title: "verify --static", ...node("scripts/verify.mjs", "--static"), groups: ["verify", ...CODE] },

  { id: "param-contract-check", title: "参数契约只读核对", ...node("scripts/param-contract.mjs", "--check"), groups: ["verify"] },
  { id: "param-contract", title: "参数契约对账（写入）", ...node("scripts/param-contract.mjs"), groups: CODE },
  { id: "param-contract-diff", title: "docs/param-contract.md 对账", ...gitDiff("docs/param-contract.md"), groups: CODE },

  { id: "spec-reproduction", title: "spec 复现（P1 验收）", ...node("test/spec-reproduction.test.mjs"), groups: ["verify", ...CODE] },
  { id: "test-static", title: "静态测试清单（scripts/lib/static-tests.mjs）", ...node("scripts/run-static-tests.mjs"), groups: CODE },
  { id: "verify-package", title: "打包冒烟 verify:package", cmd: process.env.ComSpec || "cmd.exe", args: ["/d", "/s", "/c", "npm run verify:package --silent"], groups: ["package", "release"] },

  { id: "numbers-check", title: "文档数字对账（只读）", ...node("scripts/gen-numbers.mjs", "--check"), groups: ALL },
];

if (requested === "list" || requested === "--list") {
  for (const group of ["docs", "verify", "code", "bridge", "package", "release"]) {
    console.log("\n" + group + ":");
    for (const step of STEPS.filter((s) => s.groups.includes(group))) console.log("  - " + step.id + "  " + step.title);
  }
  process.exit(0);
}

if (!["docs", "verify", "code", "bridge", "package", "release"].includes(GROUP)) {
  console.error("未知分组: " + requested + "（可用：docs | verify | code | bridge | package | release | ci）");
  process.exit(2);
}

const steps = STEPS.filter((s) => s.groups.includes(GROUP));
if (steps.length === 0) {
  console.error("分组 " + GROUP + " 没有步骤");
  process.exit(2);
}

console.log("GATES group=" + GROUP + " steps=" + steps.length + (GROUP === "release" ? "（= CI）" : ""));
const started = Date.now();
steps.forEach((step, index) => {
  const label = "[" + (index + 1) + "/" + steps.length + "] " + step.title;
  console.log("\n▶ " + label);
  const t0 = Date.now();
  const result = spawnSync(step.cmd, step.args || [], {
    cwd: step.cwd || ROOT,
    stdio: "inherit",
    windowsHide: true,
  });
  const secs = ((Date.now() - t0) / 1000).toFixed(2);
  if (result.error) {
    console.error("\nGATES FAILED: 无法启动 " + step.id + ": " + result.error.message);
    process.exit(1);
  }
  if (result.status !== 0) {
    console.error("\nGATES FAILED: " + step.id + " (exit " + result.status + ", " + secs + "s)");
    process.exit(result.status || 1);
  }
  console.log("✓ " + step.id + " (" + secs + "s)");
});
console.log("\nGATES OK (group=" + GROUP + ", " + steps.length + " steps, " + ((Date.now() - started) / 1000).toFixed(1) + "s)");
