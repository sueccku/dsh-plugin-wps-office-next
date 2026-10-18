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

下面这些**不需要 WPS**，按改动类型跑对应的**门禁分组**（秒级到约 1 分钟）。命令清单只在
`scripts/run-gates.mjs` 里维护——本文只规定"改什么跑哪组"，**不要**把命令再抄进别的文档或 CI。

| 改动类型 | 跑这一组 | 覆盖 |
| --- | --- | --- |
| 只改文档 | `npm run gates:docs` | lint（含 Markdown 链接检查）+ 文档数字对账（实测约 1 秒） |
| 只读核对（**不写工作区**） | `npm run gates:verify` | lint / `verify --static` / 参数契约只读 / spec 复现 / 技能参考表只读 / 静态测试清单 / 数字对账 |
| 改 TS / 工具 / spec | `npm run gates:code` | tsc + `mcp/dist` 对账、三个生成器 + 对账、lint、verify、参数契约（写入 + 对账）、spec 复现、`test:static`、数字对账 |
| 改桥 `mcp/scripts/wps-com.ps1` | `npm run gates:bridge` | `gates:code` + 重建 `host/wps-actions.ps1` + 对账 |
| 改 `package.json` / `files` | `npm run gates:package` | `gates:code` + `verify:package`（打包冒烟） |
| push / 发版前 | `npm run gates:release` | 全部（CI 跑的就是这一组） |

补充：
- **`verify:package` 经 `npm run` 执行**（不要直接 `node scripts\verify-package.mjs`）：两者环境不同，历史上藏过一个真缺陷（FIXES 85）。
- 门禁里的生成器会**重写**已入库的生成物；跑完 `git diff` 里出现的生成物要一起提交（否则 CI 对账会红）。
- `gates:verify` 是唯一保证"不写仓库"的分组，适合新对话开工前的只读核对。

各分组的步骤、顺序与实现只在 `scripts/run-gates.mjs` 里维护；契约管线的 7 步手工顺序见 `docs/HANDOFF.md` §4。

## 4. 发版

发版（`npm publish` / 打 tag）同样适用 §2：**不再默认跑真机整轮**，改为发版前先问用户；
用户未答复 → 不跑、不发布。（这一条推翻 R9。）

## 5. CI

`.github/workflows/ci.yml` 只做一件事：push / PR 时跑 `npm run gates:ci`（= `gates:release` 那一组，
不需要 WPS）。CI 由 GitHub 触发，不受 §2 的"先问"约束。
