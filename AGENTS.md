# 项目工作规则（本仓库的唯一权威）

> 适用目录：`D:\dsh\a`（dsh-plugin-wps-office-next）。
> 本文件是**测试时机**的唯一权威来源；README / docs 里任何相反或过期的说法，一律以本文件为准。
> 依据：用户 2026-10-10 决策（HANDOFF 决策表 **R13**），**取代 R9**「真机整轮回归 = 发版前必跑」。

## 1. 普通修改后：不跑测试

改代码 / 改文档 / 改生成物之后，**不要**主动执行完整测试或回归测试，包括：

- `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1`（真机整轮，约 15 分钟）
- `node scripts\e2e.mjs --profile <name>`（一键端到端，约 2–4 分钟）
- `node scripts\accept-install.mjs`
- 逐文件跑 `test\*.test.mjs`（`spec-reproduction.test.mjs` 例外：它不需要 WPS，属 §3 的静态门禁）

原因：真机整轮成本高；用户要求按需执行。

## 2. 要 push 到 GitHub 之前：先问用户

- 每次准备 `git push`（含发版推送）**之前**，主动问一句：**「要不要跑测试？」**，并给出范围选项（整轮 / 只跑相关文件 / 跳过）。
- **用户明确说「跑」才跑**；用户没回复 → 不跑测试，也**先不 push**，等用户回话。
- 用户说「跑」但没指定范围时，默认跑 `scripts\run-tests.ps1` 整轮。
- 本地 `git commit` 不受此约束（可随时提交）。

## 3. 照常执行的静态门禁（不属于"完整/回归测试"）

下面这些**不需要 WPS**、秒级到约 1 分钟，按老规矩照常跑（改了对应东西就要跑）：

- `cd mcp; npx tsc`（构建产物要与源码一致）
- `powershell -NoProfile -File scripts\build-host-actions.ps1`（桥源码 → 宿主生成物）
- `node scripts\gen-tool-surface.mjs`、`node scripts\gen-skill-tools.mjs`、`node scripts\gen-tool-coverage.mjs`
- `node scripts\gen-numbers.mjs --check`
- `npm run lint`
- `node scripts\verify.mjs --static`
- `node scripts\param-contract.mjs`
- `node test\spec-reproduction.test.mjs`（CI 的静态测试之一）
- `npm run verify:package`（打包冒烟：装一遍 + 启动 MCP server 握手，不需要 WPS）
- `npm run test:static`（= `node scripts/run-static-tests.mjs`；清单的唯一来源是 `scripts/lib/static-tests.mjs`）

完整的"改完代码复跑顺序"仍见 `README.md` 的「质量门禁与 CI」一节。

## 4. 发版

发版（`npm publish` / 打 tag）同样适用 §2：**不再默认跑真机整轮**，改为发版前先问用户；
用户未答复 → 不跑、不发布。（这一条推翻 R9。）

## 5. CI

`.github/workflows/ci.yml` 保持现状：push / PR 时自动跑**不需要 WPS** 的静态部分。
CI 由 GitHub 触发，不受 §2 的"先问"约束。
