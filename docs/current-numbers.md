# 当前数字（自动生成，勿手改）

> 由 `node scripts/gen-numbers.mjs` 生成；CI 跑 `--check` 核对文档里的声明值。
> 本文件只放**从干净检出就能复现**的数字。需要「跑测试才知道」的断言数**不在本文件** ——
> 它随本机是否跑过整轮而变，只写进 `test/numbers-report.json`（本机产物，.gitignore，CI 不核对）。

| 项 | 值 | 来源 |
| --- | --- | --- |
| 广告面工具 | **95 / 270** | spec/advertised.json + spec/tool-definitions.json |
| 广告面 schema 字节 | **55,510（≈15,860 tokens）** | 真实 tools/list 载荷 |
| minimal 档 | **4 工具 / 1,609 字节** | 同上 |
| full 档 | **270 工具 / 165,422 字节** | 同上 |
| 桥 action | **265** | mcp/scripts/wps-com.ps1 的 case 块（与 verify.mjs 同口径） |
| 其中带键表 | **264** | host/wps-actions.ps1 键表 |
| 桥参数键 | **1008** | 同上 |
| 带键表的 action | **263** | spec/action-keys.json |
| 未工具化 action | **2（getActivePresentation / getActiveWorkbook）** | 键表 − spec |
| 测试文件 | **64** | test/*.test.mjs |

> 「测试文件」数的是目录里现有的 `test/*.test.mjs`；README / HANDOFF 写的「N 个文件」是
> **最近一次整轮实际跑到的**数 —— 新增文件但还没跑整轮时，两者本来就该不同。

## 声明值核对

| 位置 | 文档里写的 | 实际 | 结果 |
| --- | --- | --- | --- |
| README badge advertised | 95 | 95 | ✅ |
| README badge registered | 270 | 270 | ✅ |
| README standard tools | 95 | 95 | ✅ |
| README standard bytes | 55510 | 55510 | ✅ |
| README standard tokens | 15860 | 15860 | ✅ |
| README minimal tools | 4 | 4 | ✅ |
| README minimal bytes | 1609 | 1609 | ✅ |
| README full tools | 270 | 270 | ✅ |
| README full bytes | 165422 | 165422 | ✅ |
| README install guide advertisedTools | 95 | 95 | ✅ |
| README status sample advertisedTools | 95 | 95 | ✅ |
| README detail summary | 95 | 95 | ✅ |
| README registered catalogue | 270 | 270 | ✅ |
| README current advertised | 95 | 95 | ✅ |
| README current bytes | 55510 | 55510 | ✅ |
| HANDOFF advertised | 95 | 95 | ✅ |

> 与测试计数有关的那几条声明只在**本机存在 `test/summary.json`** 时核对，**不属于 CI 核对范围**；
> 现值看 `test/summary.json` / `test/numbers-report.json`（跑完整轮后按 README / HANDOFF 的口径刷新）。
