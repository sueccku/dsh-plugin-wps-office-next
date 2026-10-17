# 兼容性与已知限制

只写**有位置、有条件**的异常。没有复现条件的传闻不要写在这里。

## 1. WPS 版本

- 版本前置检查读的是 **exe 文件版本**（`fileVersion`），不是 `Application.Version`：后者是 Office 兼容值（WPS 12.1 会报 12.0），拿它跟版本下限比较会误判。
- `wps_status` 返回 `wpsVersion` / `officeCompatVersion` / `wpsBuild` / `wpsVersionSupported` / `compatibilityWarning`。`wpsVersionSupported` 为 `false` 时先看 `compatibilityWarning`。
- 兼容性判定逻辑是 `mcp/src/utils/wps-version.ts` 里的纯函数（有单测），`wps_status` 只负责读与报。

## 2. 宿主进程与超时

- 常驻宿主是 **Windows PowerShell 5.1**（绝对路径固定为 `System32\WindowsPowerShell\v1.0\powershell.exe`）。宿主启动时会自检 `$PSMajor`，不是 5 就拒绝启动——这条不是可配置项。
- 单次调用超时：默认 60s；长动作（转换、导出、透视表、批量插入等）300s；**超时后**下一次调用先用 15s 的短超时快速失败，成功一次恢复正常。
- 超时后宿主进程被杀，**WPS 本身不会被关闭**（可能正卡在模态对话框，也可能有你未保存的内容）。状态是「未知」，不是「失败」。
- 超时后不要盲目重放非幂等操作（见 `execution-contract.md` §4）。

## 3. 单实例租约

- 同一时间只允许一个宿主操作 WPS（命名互斥体 `Local\dsh-plugin-wps-office-next-com-host` + `%USERPROFILE%\.wps-office-mcp\com-host.json`）。
- 记录：`hostPid`、`clientPid`、`startedUtc`、`updatedUtc`、`phase`、`lastAction`。
- 接管条件：**身份核对**判定原宿主已死（进程不在，或进程号被无关进程复用 —— 用 `startedUtc` 与进程 `StartTime` 比对），或原客户端已死。**仅心跳停更不再触发接管**（心跳只作为诊断信息）。
- 拒绝启动时的错误里会给出占用者的 pid、clientPid 与最近动作；确认没有别的会话在用，删掉 `com-host.json` 再重试。

## 4. 孤儿实例回收

- 宿主启动时回收**上一个被强杀的会话遗留的、由插件启动的** WPS 实例；回收前核对实例身份（类别 + 进程号 + 启动时刻），且**有未保存内容的一律不碰**。
- 身份读不出来的实例按"不回收"处理（宁可留下，也不误关用户的窗口）。
- 用户自己启动的 WPS 实例永远不在回收范围内。

## 5. 阻塞型对话框（已知规避）

- 打开**加密文档**会弹模态密码框，STA 里会阻塞整个宿主。桥的做法是给 `Open()` 传一个不可能等于真实密码的**哨兵密码**，让加密文件**快速报错**而不是无限等待。
- `DisplayAlerts` 管不到密码框，所以别指望关掉告警就能绕过它。
- 打开动作已经封装了这条规避；用 `wps_execute_method` 直接调 `Documents.Open` / `Workbooks.Open` 会绕过它，**不要这么做**。

## 6. 各应用的实测异常

| 位置 | 现象 | 处理 |
|---|---|---|
| 演示 | 某些构建上新建文稿的 `Slides.Count` 可能滞后（报 0 而页确实存在） | 以 `wps_ppt_get_slide_info` / `get_shapes` 的实际返回为准；记录构建版本 |
| 演示 | 对已有空白页 `set_slide_layout` 可能直接 `E_FAIL` | 要版式就在 `add_slide` 时指定 |
| 演示 | 扇形只暴露一个 adjustment，改角度会挂住 COM 宿主 | 多段占比用横向堆叠条 + 图例 |
| 表格 | 导出图片走剪贴板 | 并发或无人值守时可能失败；失败如实上报 |
| 表格 | `set_formula` 对多格区域**不做相对引用调整** | 逐行引用要逐格写或 `fill_series` |
| 文字 | 字符坐标在插入/删除后整体平移 | 每次重新 `get_paragraphs` |
| 文字 | `insert_page_numbers` 的"切到页脚故事"路径需要活动窗口 | 没有窗口时明确失败，而不是写错位置 |
| 全部 | 中文输出的编码链 | 客户端已按字节收帧 + `StringDecoder` 解码；**不需要**为了绕开编码问题而回避中文 |

## 7. 不要依赖的东西

- **不要**依赖"未广告的工具可以按全名直接调用"：模型能直接调用的只有 `tools/list` 里的名字，其余必须走 `wps_call`。
- **不要**依赖撤销/回滚：COM 改动不一定进撤销栈。需要可回退就先另存副本。
- **不要**依赖跨会话的目标锁定：没有 documentId，也没有跨会话句柄。
