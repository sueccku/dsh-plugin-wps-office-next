# npm 发布记录（历史）

> 本文件是**历史记录**：每次发版后的实测读数、registry 侧核对结果与事故处置。
> 现行发布流程与门禁在 [release-checklist.md](release-checklist.md)；现值数字看 [current-numbers.md](current-numbers.md)。
> 新记录追加在**最上面**（最新版本在前）。

## 0.9.0（2026-10-18）

| 项 | 读数 |
| --- | --- |
| 主题 | 四个 Skill 重写为任务决策指南 · `wps_batch` 结构化状态 · 未落盘保存关闭不再丢内容（P0）· 租约/孤儿身份收口 · 嵌套参数校验（FIXES 115）；发版前修掉新校验引出的 `format.numberFormat` 回归（FIXES 116） |
| 真机整轮 | **跑了**（用户 2026-10-18 选择「整轮」）：首轮 `TOTAL PASS=1761 FAIL=16 FILES=64`（唯一红 `cell-format`，16 红即 FIXES 116）；修复后重跑 `TOTAL PASS=1777 FAIL=0 FILES=64 REAPED=343 ORPHANS_LEFT=0`、`ALL TEST FILES GREEN`、`BASELINE OK (1777 >= 1283, 64 >= 51)` |
| 静态门禁 | `npm run gates:release` **19 步全绿**：tsc + dist 对账、宿主与四个生成器重建对账、lint（含 Markdown 链接检查）、`verify --static` 23、参数契约 259 对 0 失效、spec 复现 16、`test:static` 17/17、文档数字对账 20；`verify:package` **11/11 → `PACKAGE RUNTIME OK`** |
| 打包 | `entryCount=315  packed=766109 B (748.2KB)  unpacked=3,670,529 B (3.50MB)`；registry 实测 `fileCount=315`、`shasum=a8b5b9f55e06fc2956fed1a2325cc0cad3662db2`、`integrity=sha512-J+If3NTAGyMX2CIN0LGcfXGQYS1vfB9G4GXpZFmCYHVry60ewDnbm6WD95XvjdeKrHO6O0SKbU1UEUSBH0pM5Q==`；下载 tarball **766,109 字节**、sha1 与本地 pack 一致 |
| 全新 profile 真装 | `dsh plugin --profile wps090 add dsh-plugin-wps-office-next@0.9.0` → 装出 0.9.0、`--dump-config` 两个 id 正常、副本内 `node scripts\doctor.mjs` → `DOCTOR OK`；验收后已 `remove` + 删 profile |
| GitHub | 提交 `7cf05fb`、tag `v0.9.0`、Release <https://github.com/sueccku/dsh-plugin-wps-office-next/releases/tag/v0.9.0>；`git ls-remote --tags` 已见 `refs/tags/v0.9.0` |
| 顺带 | FIXES 116 是**发版前整轮**抓到的回归：加严嵌套未知键校验，把 schema 漏声明的 `format.numberFormat` 判成未知参数，`wps_excel_set_cell_format` 的嵌套写法整组失效；静态门禁抓不到（`arg-shape-nested` 只验 schema 内部自洽，`param-contract` 把 setCellFormat 列为 UNVALIDATED）。详见 [FIXES.md](FIXES.md) 第 116 条 |

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
| 2026-10-18 | 0.9.0 vs npm 0.8.2 | npm 0.8.2 = **294 文件**、本次 = **315 文件**，**新增 21、移除 0** —— 新增 = Skill 按需参考 11 份（6 `references` + 5 `recipes`）+ `arg-shape` / `fault-inject` 各 4 个 dist 产物 + `scripts/gen-fixes-index.mjs` / `scripts/run-gates.mjs`；**没有**漏进 `node_modules` / `test` / `src` |
| 2026-10-12 | 0.8.2 vs 0.8.1 | 两边都是 294 文件；内容相同（0.8.1 的 tarball 曾短暂 404，制品就绪后正常） |
| 2026-10-09 | 0.7.0 vs npm 0.6.3 | npm 0.6.3 = **296 文件**、本次 = **295 文件**，**新增 1、移除 2** —— 新增 `scripts/reap-wps.ps1`；移除的 `scripts/analyze-projection.mjs`、`scripts/analyze-upstream-tools.mjs` 是**本轮之前就已经删掉的死脚本**，不是打包漏件 |
| 2026-10-07 | 0.6.3 vs npm 0.6.2 | **两边都是 296 文件、新增 0、移除 0** —— 内容就地更新，没有多带（没漏进 `node_modules` / `test` / `src`）也没有少带 |
| 2026-10-04 | 0.6.1 vs npm 0.6.1 | npm = **289 文件**，本次 = **296 文件**，**新增 7 个、移除 0 个**（`mcp/dist/tools/word/style-names.{js,d.ts}` 及两个 `.map`、`scripts/gen-numbers.mjs`、`scripts/lib/tool-face.mjs`、`scripts/probe-installed.mjs`） |

> 文件数与体积**会随仓库增长漂移**，它们只是量具读数、不是断言：真正的断言在
> `verify-package.mjs`（包内必需文件 + server 真能起来），打包读数以它的 `PACKAGE SIZE` 输出为准。
