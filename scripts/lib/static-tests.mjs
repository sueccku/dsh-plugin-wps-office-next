/** **不需要 WPS** 的测试文件清单 —— 唯一来源。
 *
 * 为什么要有它：这份清单以前手写在 .github/workflows/ci.yml 的 step 里，新加一个不需要 WPS 的测试文件
 * 不会自动进 CI（门禁静默变薄）。现在 CI 用 scripts/run-static-tests.mjs 跑这份清单，
 * run-tests.ps1 -Static 也读同一份，改这里就等于改两边。
 *
 * spec-reproduction 不在这里：它在 CI 里是**具名的 P1 验收步骤**，单独跑（跑两次没有意义）。
 */
export const STATIC_TESTS = [
  "test/plugin.test.mjs",
  "test/skill-quality.test.mjs",
  "test/com-host.test.mjs",
  "test/host-lease.test.mjs",
  "test/watchdog.test.mjs",
  "test/silent-catch.test.mjs",
  "test/install-selfcheck.test.mjs",
  "test/mcp-node.test.mjs",
  "test/arg-shape-guard.test.mjs",
  "test/wps-version.test.mjs",
  "test/encrypted-preflight.test.mjs",
  "test/alerts-gate.test.mjs",
  "test/com-boundary-gate.test.mjs",
  "test/deprecated.test.mjs",
  "test/style-names.test.mjs",
  "test/fault-injection.test.mjs",
  "test/arg-shape-nested.test.mjs",
  "test/p1-surface.test.mjs",
  "test/skilleval/benchmark80.mjs",
];

/** CI 里单独跑的静态门禁（不重复进 STATIC_TESTS）。 */
export const STATIC_TESTS_SEPARATE = ["test/spec-reproduction.test.mjs"];
