# npm 发布记录（历史）

> 本文件是**历史记录**：每次发版后的实测读数、registry 侧核对结果与事故处置。
> 现行发布流程与门禁在 [release-checklist.md](release-checklist.md)；现值数字看 [current-numbers.md](current-numbers.md)。
> 新记录追加在**最上面**（最新版本在前）。

## 0.8.2（2026-10-12）

| 项 | 读数 |
| --- | --- |
| 主题 | 宿主 PATH 上没有 Node.js 时 **95 个工具全部不可用**（FIXES 114） |
| 真机整轮 | **跑了**（用户 2026-10-12 明确要求「跑测试」）：`TOTAL PASS=1508 FAIL=0 FILES=59 REAPED=240 ORPHANS_LEFT=0`、`ALL TEST FILES GREEN`、`BASELINE OK (1508 >= 1283, 59 >= 51)` |
| 静态门禁 | tsc / build-host-actions / 三个生成器 / `gen-numbers --check` / lint / `verify --static` / `param-contract` / spec-reproduction / `test:static` 全绿；`verify:package` **11/11 → `PACKAGE RUNTIME OK`** |
| 打包 | `entryCount=294  packed=706231 B  unpacked=3.35MB`；registry 实测 `fileCount=294 shasum=966ceabe3ea1293a84482fbbb24f1ed969a70fdc`，下载 tarball **706231 字节**与本地 pack 一致 |
| 全新 profile 真装 | `dsh plugin --profile <临时> add dsh-plugin-wps-office-next@0.8.2` → 装出 0.8.2、两个 id 接线正常，验收后已拆掉 |
| 顺带 | `run-tests.ps1` 的 P6 底线门禁此前**静默失效**（`test/baseline.json` 无 BOM + PS 5.1 按 ANSI 读 → `ConvertFrom-Json` 抛错、违规数组仍为空 → 放行），本版一并修掉 |

## 0.8.1（2026-10-12）

| 项 | 读数 |
| --- | --- |
| 结果 | 代码与 GitHub 侧完整可用（提交 `b2ffefe`、tag `v0.8.1`）；npm 上**最终也正常**（tarball 200、`shasum=c95c1d13867b305cfe47998eb5b07a24cc47f82b`），但与 0.8.2 内容重复 |
| 插曲 | `npm publish` 返回 **HTTP 202**（不是 200）；随后 `GET …/0.8.1` 200、`dist-tags.latest=0.8.1`，但 **tarball 约半小时 404**；期间重发同号 → `E409 Cannot publish over previously staged version "0.8.1"`；`npm stage list` 报无待批准项，`npm stage approve` 无从下手。制品就绪后 tarball 变 200 |
| 处置 | 当时据此换号发了内容相同的 0.8.2；`latest = 0.8.2` |
| 教训 | ①`npm publish` 之后必须以「**HTTP 200 + `dist.tarball` 可 HEAD**」为判据，不能只看退出码和 `+ pkg@version`；②看到 tarball 404 先等 20–30 分钟复检再决定换号，202 表示发布被接受、制品可能仍在就绪中；③`npm stage list` 在这条路径上不显示任何待批准项，别据此判断"暂存卡住了" |

## 打包逐文件对比（历史）

发布前可选做的一步：把 npm 上上一版的 tarball 解开，与本次 `npm pack` 的文件清单 diff ——
应当只剩"本次新增的文件"，且**没有移除**。命令见 [release-checklist.md](release-checklist.md) §2。

| 日期 | 对比 | 读数 |
| --- | --- | --- |
| 2026-10-12 | 0.8.2 vs 0.8.1 | 两边都是 294 文件；内容相同（0.8.1 的 tarball 曾短暂 404，制品就绪后正常） |
| 2026-10-09 | 0.7.0 vs npm 0.6.3 | npm 0.6.3 = **296 文件**、本次 = **295 文件**，**新增 1、移除 2** —— 新增 `scripts/reap-wps.ps1`；移除的 `scripts/analyze-projection.mjs`、`scripts/analyze-upstream-tools.mjs` 是**本轮之前就已经删掉的死脚本**，不是打包漏件 |
| 2026-10-07 | 0.6.3 vs npm 0.6.2 | **两边都是 296 文件、新增 0、移除 0** —— 内容就地更新，没有多带（没漏进 `node_modules` / `test` / `src`）也没有少带 |
| 2026-10-04 | 0.6.1 vs npm 0.6.1 | npm = **289 文件**，本次 = **296 文件**，**新增 7 个、移除 0 个**（`mcp/dist/tools/word/style-names.{js,d.ts}` 及两个 `.map`、`scripts/gen-numbers.mjs`、`scripts/lib/tool-face.mjs`、`scripts/probe-installed.mjs`） |

> 文件数与体积**会随仓库增长漂移**，它们只是量具读数、不是断言：真正的断言在
> `verify-package.mjs`（包内必需文件 + server 真能起来），打包读数以它的 `PACKAGE SIZE` 输出为准。
