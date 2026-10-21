# 文件级批量配方（P2-1，2026-10-21）

> 定位：**插件不提供通用文件系统工具**（列目录/改名/移动/删除由 Harness 的文件工具做）。
> 插件只负责「打开 → 操作 → 另存/转换 → 关闭」，批量编排由调用方用 `wps_batch` 或 Harness 循环完成。
> 所有批量任务都按 **O(N) 次调用** 计成本，并必须产出**失败清单**；不要因为"大部分成功"就报全绿。

## 通用骨架

1. **列清单**：用 Harness 的文件工具拿到输入文件绝对路径；不要靠 `get_open_*` 猜。
2. **先试一个**：第一份成功后再批量；第一份失败先停下修参数。
3. **逐份处理**（每份 4 步）：
   - `wps_*_open_*`（只读任务 `readOnly:true`，要改的 `readOnly:false`）；
   - 读/改/转换；
   - `wps_common_save_as`（`createDirs:true`）或 `wps_convert_to_pdf`；
   - `wps_*_close_*`（`save:false`，除非要覆盖原件）。
4. **失败清单**：每份记录 `{file, step, error, retryable}`；不要吞掉错误。
5. **验收**：产物数量 = 输入数量 − 失败数；抽样打开 2–3 份；源文件指纹未变（只读任务）。

## XA9 批量导出 PDF

```text
输入：目录下的 .docx/.xlsx/.pptx
对每份：
  open_* readOnly:true
  convert_to_pdf { appType: word|excel|ppt, outputPath: <out>/<同名>.pdf }
  close_* save:false
失败：记入 failures.json；会员/风控/加密/损坏分别标注
```

- `appType` 必须显式传，否则桥可能拿到错误的应用（历史 FIXES 16）。
- 输出目录用 `save_as`/`convert_to_pdf` 的 `createDirs:true`（P0 起）或 Harness 先建。
- 不要把 50 份塞进一次 `wps_batch` 且默认 `stopOnError:true`；批量任务显式 `stopOnError:false`，结束后按 `results[].ok` 生成失败清单。

## XA10 批量改名/整理

- **完全交给 Harness 的文件工具**；不要为了改名去 `open` 再 `save_as`（会触发 WPS 会员/重名/锁定）。
- 规则先写成 dry-run 清单（旧名 → 新名），确认无冲突再执行。
- 插件只在改名后需要重算链接/更新目录时介入。

## XA11 批量套模板/统一格式

```text
对每份：
  open_* readOnly:false
  save_as 到 <out>/<原名>（先备份，createDirs:true）
  改样式/页眉页脚/表格/图表
  save { path: 副本 } 或 save_as 覆盖副本
  close_* save:false
验收：抽样 2–3 份；对比源文件指纹；记录模板版本
```

- 改原件前必须 `save_as` 副本；默认 `open_*` 是只读，批量改文件要显式 `readOnly:false`。
- 批量模板不要用 `wps_batch` 把「open + 改 + save + close」拆成互相依赖的步骤；每一步单独调用，失败时能定位。
- 100 份以上先抽样 3 份确认模板；WPS 会员/并发限制可能让第 N 份失败，失败清单要保留原始错误。

## 失败清单格式

```json
{
  "total": 50,
  "succeeded": 47,
  "failed": 3,
  "failures": [
    { "file": "D:\\in\\a.docx", "step": "open", "error": "...", "retryable": false }
  ]
}
```

报告时分开写：**已执行 / 已核对 / 未验证**；不能把"批处理跑完了"写成"50 份都正确"。
