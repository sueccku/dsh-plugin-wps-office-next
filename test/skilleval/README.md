# Skill 评估脚手架（固定场景 + 确定性判定）

对应 [docs/skill-evaluation.md](../../docs/skill-evaluation.md) 的 15 个固定场景。
目的：让"新旧 Skill 谁更好"可以**重复测量**，而不是靠印象。

## 它能做什么 / 不能做什么

| 能做 | 不能做（别当成已通过） |
|---|---|
| 一键生成每个场景的输入夹具，并算输入指纹 | `run.mjs --dry-run` / `verify.mjs` 不跑模型；**只有 `probe.mjs` 会跑** |
| 每个场景的提示词（`prompt.txt`）与断言清单（`checks.txt`） | 还没 ready 的场景只有提示与成功标准，没有夹具/断言 |
| 对已有产物做**确定性**判定（文件/文本/公式/原文件未改/轨迹匹配） | 视觉验收（溢出、重叠）不在本脚手架范围 |
| 跑一次 headless 会话并落盘轨迹 + 指标（`probe.mjs`） | 指标里的"重复询问/tokens"目前只给了原始轨迹，聚合口径在 D1 定 |

## 用法

```powershell
# 1) 只准备：生成夹具 + 指纹 + 提示 + 断言清单（不调模型、不需要 WPS）
node test/skilleval/run.mjs --dry-run

# 单个场景；已存在的 input/ 默认**不重建**（防止把上一轮的产物 rm -rf 掉），要重建显式 --force
node test/skilleval/run.mjs --dry-run --scenario excel-row-formulas
node test/skilleval/run.mjs --dry-run --scenario excel-row-formulas --force

# 2) 看"跑一轮"的确切命令（**不执行**）
node test/skilleval/run.mjs --print-command --profile <profile>

# 3) 对产物做判定（工作目录默认是 test/.artifacts/skill-eval/<NN>-<key>/input）
node test/skilleval/verify.mjs excel-row-formulas
```

## 跑一轮的正确姿势：input/ 只读，每轮复制一份

`input/` 是**夹具**，不要在上面直接跑；每轮复制出一份再跑，多轮对比才互不覆盖：

```powershell
Copy-Item -Recurse -Force "<NN>-<key>\input" "<NN>-<key>\run-r1"
cd "<NN>-<key>\run-r1"
dsh --profile <profile> "<prompt.txt 里的提示>"
node test\skilleval\verify.mjs <key> --dir "<NN>-<key>\run-r1"
```

`--print-command` 打印的就是这套命令（含复制与判定两行）。

## 跑一次 headless 会话并取证（`probe.mjs`）

`probe.mjs` 是"自由提示词"的探针：跑一次 `dsh --profile <p> <task>`，然后**解码会话日志**
（`$DSH_HOME/sessions/*/session.v*.jsonl.zstd`，按 zstd 帧切分）把工具调用、工具结果、助手文本落盘。
P1 的 §3.2 资源读取验证就是用它做的；D1 的指标也从它的 `metrics.json` 取。

```powershell
# 需要模型凭据（headless profile 从环境读；web/desktop 端存在 credentials 服务里）
$env:DEEPSEEK_API_KEY = <key>
node test\skilleval\probe.mjs --profile <profile> --label s32 --dir <工作目录> --prompt <提示词文件>
```

产出 `<工作目录>/probe-<label>/`：

```
prompt.txt     我们发给它的任务原文
stdout/stderr  dsh 的进程输出（stderr 里是模型的推理流）
transcript.txt 重排后的 CALL / RESULT / ASST 行
metrics.json   工具调用与结果全文 + 计数（toolCalls / skillLoads / helpQueries / finalAnswer）
session.txt    这次运行对应的会话文件路径
```

⚠️ 它会花真钱、并可能让 Agent 改动工作目录里的文件。**本仓库的规矩是跑这类会话前先问用户**
（`AGENTS.md` §1）。脚本自己也只做一件事，不检查你有没有获得同意。
判 dsh 启动器时不会直接 `spawn('dsh')`（Windows 上是 `.cmd`，Node 不带 shell 会 ENOENT），
而是解析到真正的 `lib/bin.js` 用当前 node 跑，见 [dsh-bin.mjs](dsh-bin.mjs)。

产物落在 `test/.artifacts/skill-eval/<NN>-<key>/`（**已被 .gitignore 忽略**）：

```
01-word-new-save/
  input/             ← 被测 Agent 的工作目录（夹具在这里）
  fingerprint.json   ← 输入指纹（unchanged: 断言用它）
  prompt.txt         ← 固定提示词
  checks.txt         ← 断言清单 + 成功标准
```

## 跑 N 轮并汇总（`rounds.mjs`，D1 的执行器）

```powershell
$env:DEEPSEEK_API_KEY = <key>
node test\skilleval\rounds.mjs --profile <profile> --label new --rounds 3 [--scenario <key|all>] [--reap-wps] [--resume]
```

- 每轮：清场（可选）→ 复制 `input/` 成 `run-<label>-r<N>/` → headless 跑一次 → 解轨迹 → `verify.mjs` 判定；
  结果写 `run-<label>-r<N>/result.json`，汇总写 `_summary/rounds-<label>.{json,md}`。
- `--resume`：跳过已有 `result.json` 的轮次，中断后接着跑。
- `--reap-wps`：每轮开始前跑 `scripts/reap-wps.ps1 -Since <本臂最早一轮的时刻>`，把评测自己留下的 WPS 清掉
  （否则第二、三轮跑的是"WPS 已经开着"，场景 1 的冷启动前提就没了）。**它真的会杀进程**，只在测试机上用。
- **单轮默认超时 60s**（`--timeout` 可改）：正常一轮 20–80s；超时的轮次记 `timedOut: true` 并判失败，
  同时用 `taskkill /T` 收掉整棵进程树 —— 不能让孙进程（MCP server / COM 宿主）活下来占着 WPS 租约。
  个别慢场景可以在 `scenarios.mjs` 里用 `timeoutSec` 单独放宽（目前只有 #14 partial-failure，实测要 90–120s）。
- `--reverify`：不跑模型，只对已落盘的运行目录重算 verify（改了断言后省一次模型钱）。
  **注意**：夹具重新生成后 `unchanged` 断言会失效（docx/xlsx 的 zip 时间戳让同一个夹具两次生成字节不同），
  所以只重建**缺失**的夹具（`--dry-run` 不带 `--force`）；一旦 `--force` 重建了某个场景，
  该场景的旧轮次必须重跑，`--reverify` 会跳过而不是把它们改判成 FAIL。
- 每轮记录：`ok`、`timedOut`、退出码、耗时、**真实 token**（input/output/cacheRead/total）、工具调用序列、
  `usedWpsTools`（有没有真的走 WPS 通道而不是用 python 绕过）、`verify` 的逐条断言。

## A/B 对照怎么搭（旧 Skill 基线）

新旧两条臂要**同一 profile 模板、同一夹具、同一提示词**，只换插件代码。旧臂用 git worktree 固定到重构前的提交：

```powershell
git worktree add D:\dsh\baseline-skill <重构前的 commit>
cmd /c mklink /J "D:\dsh\baseline-skill\mcp\node_modules" "D:\dsh\a\mcp\node_modules"   # mcp/dist 需要运行时依赖
dsh --profile wpse2e-old --from-default-profile headless --dump-config
dsh plugin --profile wpse2e-old add D:\dsh\baseline-skill
# 校验：两个 profile 的 --dump-config 都有 wps-office-next-plugin / mcp-wps-office-next
node test\skilleval\rounds.mjs --profile wpse2e-old --label old --rounds 3 --reap-wps
```

两臂**必须串行**跑（WPS 是单实例，租约会让并发的第二个宿主拒绝启动）。

跑完两臂后合并成一张对比表：

```powershell
node test\skilleval\compare.mjs --new new --old old
```

它会读两臂的 `_summary/rounds-<label>.json`，按场景对齐输出成功率差、平均 token 差、失败断言，
写 `_summary/compare.md`（本机产物）。其中「无依据成功声明」是**启发式**代理指标
（判定失败但最终回答出现"已完成/已保存"等措辞），只能提示过度自信，不是逐条审计。

## Python 依赖

夹具生成与产物检查用 `openpyxl` / `python-docx` / `python-pptx`。脚本按这个顺序找解释器：
`--python <exe>` → 环境变量 `SKILLEVAL_PYTHON` → `$DSH_HOME/dsh-runtimes/*/dependencies/python/python.exe` → PATH 上的 `python`。
找不到（或库不全）会直接报错并给出修法，不会退化成"空结果当通过"。

## 断言语法

见 [verify.mjs](verify.mjs) 顶部注释：

| 断言 | 说明 |
|---|---|
| `file-exists` / `file-absent` | 文件在 / 不在 |
| `docx-text:<rel>\|<子串>` | docx 正文+表格文本包含子串 |
| `docx-xml-count:<rel>\|<期望个数>\|<正则>` | word/document.xml 里正则命中几次（数 TOC 域、查节属性） |
| `docx-section-count:<rel>\|<期望>` / `docx-section-orientation:<rel>\|<1基节号>\|<portrait\|landscape>` | 节数与逐节方向（python-docx 按 pgSz 宽高判） |
| `xlsx-value` / `xlsx-formula:<rel>\|<sheet>\|<cell>\|<期望>` | 单元格值 / 公式文本 |
| `pptx-slide-count:<rel>\|<期望>` / `pptx-text:<rel>\|<1基页号>\|<子串>` | 页数 / 某页文字 |
| `pptx-slide-xml-matches:<rel>\|<1基页号>\|<正则>` | 某页 slideN.xml 匹配（查 `<p:timing>` 动画之类） |
| `unchanged:<rel>` | 与 `fingerprint.json` 一致（原文件未被改） |
| `transcript-matches:<正则>` / `transcript-not-matches:<正则>` | 会话轨迹匹配 / 不匹配 |

`unchanged` 需要 `fingerprint.json`；轨迹类断言需要同目录有 `transcript.txt` / `run.stdout.txt` / `trace.txt`，
缺了记 `SKIP` 而不是 `PASS`。

## 受控故障注入（仅测试用）

超时与"部分未生效"在真机上没法稳定复现：调小全局超时会让整轮所有调用一起坏，测出来是"环境坏了"，
不是"模型怎么处理失败"。所以插件里有一个**默认完全惰性**的注入开关（`mcp/src/utils/fault-inject.ts`），
只有测试运行器给子进程设了 `WPS_OFFICE_FAULT_INJECT` 才生效：

```
WPS_OFFICE_FAULT_INJECT="<tool正则>:<mode>[:<count>];..."      # 多条用 ; 分隔，count 默认 1
  mode=timeout     执行完动作后如实回「状态未知」超时（动作其实已落盘）→ 测"超时后会不会盲目重放"
  mode=shortfall   不执行该动作，直接按"这一步未生效"回报（success:false + shortfall）→ 测"能不能如实说没做到"
```

- 匹配 tool 名；`wps_call` 还会用内层 tool 名和 `args.method/action`（防"绕行调桥动作"绕过注入）。
- 不设这个变量时：不解析、不改结果、不产生日志，正常路径行为与没有这个开关时一致（`test/fault-injection.test.mjs` 钉住）。
- 用它的场景只有两个：#12 `write-timeout`（`INSERT_TEXT:timeout:1`）与 #14 `partial-failure`
  （`set_font/apply_style:shortfall:99`，并单独放宽到 120s——持续 shortfall 下模型要反复排查，实测 90–120s）。

## 为什么夹具不入库、代码入库

夹具是机器本地产物（每个场景的输入文件几百 KB 到几 MB，且要能按需重建）；
**场景表、提示词、断言、夹具生成脚本**入库，保证同一份场景在任何机器上可重建、可复核。

## 场景就绪状态

15 个场景里 **14 个 ready**（有夹具 + 确定性断言），1 个 `ready: false`：

| # | 场景 | 为什么不做 |
|---|---|---|
| 2 | only-word-skill | 不是任务场景；插件目前无条件注册四个技能，没有"只注册一个"的配置项。它要回答的问题已由 §3.2 实测（应用 Skill 单独加载**不会**带出总 Skill） |

#12/#14 原先也在此列，现在用上面的注入开关补齐了。未跑的场景**不计入**"已跑"。
