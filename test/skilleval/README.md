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

## Python 依赖

夹具生成与产物检查用 `openpyxl` / `python-docx` / `python-pptx`。脚本按这个顺序找解释器：
`--python <exe>` → 环境变量 `SKILLEVAL_PYTHON` → `$DSH_HOME/dsh-runtimes/*/dependencies/python/python.exe` → PATH 上的 `python`。
找不到（或库不全）会直接报错并给出修法，不会退化成"空结果当通过"。

## 断言语法

见 [verify.mjs](verify.mjs) 顶部注释：`file-exists` / `file-absent` / `docx-text` / `xlsx-value` /
`xlsx-formula` / `unchanged` / `transcript-matches` / `transcript-not-matches`。
`unchanged` 需要 `fingerprint.json`；轨迹类断言需要同目录有 `transcript.txt` / `run.stdout.txt` / `trace.txt`，
缺了记 `SKIP` 而不是 `PASS`。

## 为什么夹具不入库、代码入库

夹具是机器本地产物（每个场景的输入文件几百 KB 到几 MB，且要能按需重建）；
**场景表、提示词、断言、夹具生成脚本**入库，保证同一份场景在任何机器上可重建、可复核。

## 还没 ready 的场景

`scenarios.mjs` 里 `ready: false` 的 9 个场景（跨工作簿、目录刷新、多节、PPT 页序、超时、部分失败等）
需要更复杂的夹具或受控故障注入，按 [skill-evaluation.md](../../docs/skill-evaluation.md) §3 的说明在 D1 补。
它们**不计入**"已跑"。
