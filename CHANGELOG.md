# 更新日志

本文件记录每个发布版本的用户可见变化；逐条修复的原因与实测证据见 [docs/FIXES.md](docs/FIXES.md)。

## 0.3.0（稳定性加固第 1 波：不再卡死、不再撞车）

功能面与 0.2.1 完全一致：工具数、广告面（69 工具 / 37,573 字节）、参数契约、桥 action 都没有任何改动。
这一版只做**稳定性**——目标从「再加能力」转为「让客户敢用、出问题能解释、坏不了数据」。

### 亮点

- **不再被密码框卡死**：打开带密码的文件时 WPS 会弹一个模式对话框，此前会让整个 DSH 会话**永久卡住**
  （杀宿主也没用，框还在）。现在改成**不弹框、1 秒内直接报错**，并告诉用户怎么办。
- **同一时间只允许一个宿主**：同时开两个 DSH 会话操作 WPS 时，第二个会话会收到一句**可读的中文错误**，
  而不是两边互相卡死。
- **超时不再等于「已失败」**：超时文案明确说明「状态未知（可能被对话框阻塞）」，并且**不会自动关闭 WPS**
  （那会丢掉你没保存的内容）；之后会先用短超时快速失败，**成功一次即恢复**。
- **修掉两个恢复路径缺陷**（由本轮新测试抓出）：陈旧子进程的 `exit` 会反杀刚拉起的新宿主；
  `ready` 帧误清 `suspect` 标志，使短超时形同虚设。

### 实测依据

`docs/FIXES.md` 第 52 条记录了完整探针数据。两处计划假设被实测推翻：`PasswordDocument:=""` **挡不住**
密码框（空串等于没给密码）；扩展名不符的老 `.doc` 实测**不弹**转换框。真正有效的是传**非空哨兵密码**——
加密 `.docx` 从「永久卡死」变成 4 秒内报错，加密 `.xlsx` 1 秒内报错（`0xFFF40006`）。
另一处偏离计划：超时后**不杀整个 WPS 进程树**，因为实测本机有 243 个进程叫 `wps`、且 Word/PPT 的 COM
对象不提供 `Hwnd`，「被卡住的是哪一个」无法可靠指认，照做会连带杀掉用户整套 WPS。

### 已知限制（如实记录）

- 加密的**演示文稿（.pptx）**是唯一没堵住的口子：`Presentations.Open` 没有密码参数，无法阻止弹框；
  请先用 WPS 手工打开它并另存为不加密副本，再对副本操作。
- `wps_execute_method` 逃生舱下，`Document.Password` 按长度失效：15 字符卡死调用，**17 字符静默写出
  完全不加密的文件**。本插件的正式工具里没有「设置打开密码」的能力。

### 验收

`test/open-safety.test.mjs` 21 项、`test/host-lease.test.mjs` 18 项、`test/watchdog.test.mjs` 15 项全绿；
全套 **595 项 / 28 文件** + `verify.mjs` 23 项。后两个测试不需要 WPS，已并入 GitHub Actions 静态门禁。

### 安装

```powershell
dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.3.0
```

连不上 github.com 时改用：

```powershell
dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.3.0
```

## 0.2.1（仅文档与打包修正，无行为变化）

功能与 0.2.0 完全一致：工具数、广告面、参数契约、桥 action 都没有任何改动。这一版只修文档与打包。

### 改动

- **README 重写为面向使用者的两段式**：新增「它适合谁 / 能帮你做什么 / 你可以这样提要求」，安装改成
  「把一段话复制给你的 DSH AI」+ 一节写给 AI 的 7 步安装指引（环境要求、判断 profile、执行安装、
  确认接线、重启验证、排错表、环境自检），并补上卸载、重要注意事项与已知限制；技术细节收进折叠区。
- **`CHANGELOG.md` 现在会随包安装**（此前 `files` 白名单漏了它，0.2.0 的安装包里没有这个文件）。
  `README.md`、`LICENSE`、`THIRD_PARTY_NOTICES.md` 原本就在白名单里，现在文档齐了。
- **安装指引逐条实测过**：profile 探测脚本、codeload 备用地址、`--dump-config` 的接线检查、
  `scripts/doctor.mjs`、`remove` 子命令、批量上限 50——全部与本仓库实际行为一致。
- **修掉一处会把人带偏的提示**：`dsh plugin add` 失败时，dsh 会补一句「构建脚本被拦截 / 请加 allowBuilds」
  的通用兜底文案；本包没有 `prepare` 脚本、也没有原生依赖，**永远不需要改 allowBuilds**。
  真因通常是这台机器连不上 github.com，用 codeload 地址重试即可（README 已写明）。

### 安装

```powershell
dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.2.1
```

连不上 github.com 时改用：

```powershell
dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.2.1
```

## 0.2.0（首个发布）

把 [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) 与
[CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) 合并成**一个 DSH bundle**：
DSH 插件 + 自带 MCP server + 常驻 COM 宿主 + 全套技能。仅 Windows、仅 COM，不需要任何 WPS 加载项。

**广告面比上游基线小 73.5%，能力反而更全**：250 工具 / 141,872 schema 字节 → **69 工具 / 37,573 字节**；
注册工具 267 个（其余经 `wps_call` 与技能参考表触达）；541 项测试全部驱动**真实 WPS**。

### 亮点

- **常驻 COM 宿主**：取代「每次新起一个 PowerShell」，冷启动约 1.0s、稳态 1–2ms，支持串行队列与崩溃重启
- **契约真源**：工具 schema、桥的键表、技能参考表都由一份操作规格生成，CI 断言「源码与产物逐字节一致」
- **Excel 做深**：表（ListObject）全族、条件格式与数据验证的读与删、页面设置与打印（打印标题/页眉页脚/横向 A4）、
  单变量求解、迷你图、透视表刷新与清除、公式审计（引用与被引用）、命名范围读写
- **Word 做深**：表格读写编辑（增删行列/合并拆分/样式）、页码、分栏、修订列表与接受拒绝、批注读删、
  内容控件、脚注尾注、索引、交叉引用、**邮件合并**（CSV → 生成新文档，母版不动）
- **PPT 做减法**：放弃 3D 族与美化族，把碎片 setter 归并（形状效果四合一、动画三合一、表格样式三合一、
  页脚三合一），88 → 76 个工具
- **一键 e2e**：一条命令造 fixture、跑真实 headless 任务、逐帧解会话日志、用裸 COM 重开产物核对，28 项检查

### 修好的「从未生效」缺陷

这些能力早就写在桥里，却因为宿主环境差异或参数写法而**永远不工作**，此前没有任何测试覆盖：

1. **查找**：`Range.Find()` 结果的 `Address()` 在常驻宿主里不可用 → 改为按 A1 串取范围、一次 `Value2` 扫描
2. **二维数组下标**：PowerShell 的逗号优先级高于 `+`，`$m[$a + $r, $b + $c]` 被解析成
   `$a + ($r, $b) + $c`，异常又被裸 `catch { continue }` 吞成「0 命中」→ 下标加括号
3. **合并计算 / 分类汇总**：`XlConsolidationFunction` 常量写成了假值，每次调用都 HRESULT 0x800A03EC
4. **表（ListObject）**：`Delete()`/`Add()` 之后同一个对象仍返回旧几何，工具会把变更前的结构报给用户
5. **透视表**：`Get-RangeFromAddress` 的 `return $range` 被 PowerShell 展开成 21 个单格，创建透视表**从来没有成功过**
6. **脚注之后**：光标留在注释正文里，此后所有「往光标处插内容」的动作都插进了注释（邮件合并因此静默出错）

### 实测的 WPS 限制（不做，如实记录）

- **水印**：WPS 的页眉 `Shapes` 集合不接受任何图形（`AddTextEffect`/`AddShape`/`AddTextbox` 都返回对象但 `Count` 恒为 0）
- **文档属性**：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳（`.Item()`、`GetType()` 直接抛 null 引用）
- **切片器**：`SlicerCaches.Add2` 建得出缓存，但 `Slicers.Count` 恒为 0，用户可见的切片器不会出现
- **场景管理器**：`Worksheet.Scenarios` 在 COM 里被暴露成方法，语义读不干净

以上四项都可以用隐藏逃生舱 `wps_execute_method` 自行尝试（默认不在广告面里）。

### 安装

```powershell
dsh plugin --profile <name> add github:sueccku/dsh-plugin-wps-office-next#v0.2.0
```

要求：Windows、WPS Office 12.1+ x64。预构建产物已入库，安装后开箱可用，无需构建步骤。
