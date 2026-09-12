# host/

常驻 WPS COM 宿主。替代「每次工具调用 spawn 一个 PowerShell 并重新解析 4,900 行脚本」的旧路径。

## 文件

| 文件 | 说明 |
|---|---|
| wps-actions.ps1 | 生成物，勿手改。由 scripts/build-host-actions.ps1 从 mcp/scripts/wps-com.ps1 转换而来 |
| wps-com-host.ps1 | 宿主主循环，手写并维护 |

## 协议

stdin 每行一个 JSON 请求，stdout 每行一个 JSON 响应，UTF-8：

```
请求  {"id": 1, "action": "getCellValue", "params": {"sheet": "Sheet1", "row": 1, "col": 1}}
响应  {"id": 1, "ok": true, "result": {"success": true, "data": {...}}, "ms": 2}
```

启动后宿主先输出一行 ready 帧 `{"ready":true,"pid":...,"protocol":1}`，再进入循环。
`action` 为 `__shutdown` 时正常退出。

## 关键约束

- 仅 stdout 承载协议帧；action 调用的成功输出流被丢弃，错误流交给父进程采集，避免污染协议。
- 结果不再走 stdout，而是写入 `$script:WpsResult`，因此宿主可以安全地 `$null = Invoke-WpsAction ...`。
- 必须用 `powershell.exe` 5.1（支持 `-STA`），不要换成 pwsh。
- wps-actions.ps1 必须带 UTF-8 BOM，否则 PowerShell 5.1 会把中文当 GBK 读。`.gitattributes` 的 `* -text` 与生成脚本都会保证这一点。
- 上游 367 处 `exit` 已转换为 `return`；重新生成后可检查 exit_remaining 是否为 0。

## 重新生成

```
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build-host-actions.ps1
```

生成脚本会在 param 块、dispatch 边界或 Output-Json 形态变化时直接报错，而不是静默产出坏文件。
