# 当前数字（自动生成，勿手改）

> 由 `node scripts/gen-numbers.mjs` 生成；CI 跑 `--check` 核对文档里的声明值。
> 需要「跑测试才知道」的数字（断言数）来自 `test/summary.json`（由 `scripts/run-tests.ps1` 每次整轮写出，不入库）。

| 项 | 值 | 来源 |
| --- | --- | --- |
| 广告面工具 | **84 / 268** | spec/advertised.json + spec/tool-definitions.json |
| 广告面 schema 字节 | **46,867（≈13,391 tokens）** | 真实 tools/list 载荷 |
| minimal 档 | **4 工具 / 1,353 字节** | 同上 |
| full 档 | **268 工具 / 157,854 字节** | 同上 |
| 桥 action | **263** | mcp/scripts/wps-com.ps1 的 case 块（与 verify.mjs 同口径） |
| 其中带键表 | **262** | host/wps-actions.ps1 键表 |
| 桥参数键 | **1000** | 同上 |
| 带键表的 action | **261** | spec/action-keys.json |
| 未工具化 action | **2（getActivePresentation / getActiveWorkbook）** | 键表 − spec |
| 测试文件 | **48** | test/*.test.mjs |
| 测试断言 | **1077（失败 1）** | test/summary.json（该行不含跑测时间戳：会让 CI 对账误报） |

## 声明值核对

| 位置 | 文档里写的 | 实际 | 结果 |
| --- | --- | --- | --- |
| README badge advertised | 84 | 84 | ✅ |
| README badge registered | 268 | 268 | ✅ |
| README standard tools | 84 | 84 | ✅ |
| README standard bytes | 46867 | 46867 | ✅ |
| README standard tokens | 13391 | 13391 | ✅ |
| README minimal tools | 4 | 4 | ✅ |
| README minimal bytes | 1353 | 1353 | ✅ |
| README full tools | 268 | 268 | ✅ |
| README full bytes | 157854 | 157854 | ✅ |
| README install guide advertisedTools | 84 | 84 | ✅ |
| README status sample advertisedTools | 84 | 84 | ✅ |
| README detail summary | 84 | 84 | ✅ |
| README registered catalogue | 268 | 268 | ✅ |
| README current advertised | 84 | 84 | ✅ |
| README current bytes | 46867 | 46867 | ✅ |
| README test assertions | 1077 | 1077 | ✅ |
| README test files | 48 | 48 | ✅ |
| HANDOFF test assertions | 1077 | 1077 | ✅ |
| HANDOFF test files | 48 | 48 | ✅ |
| HANDOFF advertised | 84 | 84 | ✅ |

