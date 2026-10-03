# npm 发布检查单（release checklist）

> 适用：把本仓库作为 **npm 公开包** `dsh-plugin-wps-office-next` 发布到 `https://registry.npmjs.org/`。
> 决策依据：D18 由 B 改为 A（开始发布）、D2 = A（`files` 改成显式子路径）、D3 = A（首发版本 0.6.0）、
> D4 = B（`publishConfig.registry` 入库、凭证不入库）、D5 = A（只维护这份文档检查单）。
> 这份文档是**手动流程**：本仓库没有发布用的 CI 工作流。
>
> **首次真发布已发生**：`dsh-plugin-wps-office-next@0.6.0` 于 **2026-10-02** 发布到 registry.npmjs.org。
> **2026-10-02 实测**：`npm login` 之后 `npm publish` 仍会报 `EOTP`，需要按提示做一次浏览器授权（本机就是这样过的）；
> 想彻底免掉 OTP，得换成勾了 `Bypass 2FA` 的 Granular token（或 Automation 类型的 classic token）。

---

## 0. 一次性准备

| 项 | 要求 | 怎么确认 |
| --- | --- | --- |
| npm 账号 | 已注册并登录 registry.npmjs.org | `npm whoami --registry https://registry.npmjs.org` 打印出用户名 |
| 发布凭证 | 推荐 **Granular Access Token**：权限 Read and write、勾选 **Bypass 2FA**、有效期自定，只对本包有效 | 写入用户级 `~/.npmrc`：`//registry.npmjs.org/:_authToken=<token>` |
| 或者 | 发布时按 npm 提示做一次浏览器授权 / 输 6 位 OTP | 见下面的「2026-10-02 实测」 |
| 包名归属 | 首发之后，包的 owner 就是发布账号 | `npm owner ls dsh-plugin-wps-office-next` |
| 首次发布权限 | npm 对**从未发布过的包**默认算 public；若报 `402 Payment Required`，加 `--access public` | — |

**2026-10-02 实测（本机真实遇到的两件事）**

1. 本机原先 `registry=https://registry.npmmirror.com/`（只读镜像，**不接受发布**）。为发布做了 `npm login` 后，
   `~/.npmrc` 的 `registry` 被改写成 `https://registry.npmjs.org/`，并多出一行
   `//registry.npmjs.org/:_authToken=…`（token 只在本机，**没有进仓库**）。副作用：这台机器之后所有 `npm install`
   都从 npmjs 拉包，不再走镜像 —— 需要镜像的话自己把 registry 那行改回去。
2. 这个 token **仍然会触发 OTP**：直接发布时报

   ```
   npm error code EOTP
   npm error This operation requires a one-time password.
   npm error Open this URL in your browser to authenticate: https://www.npmjs.com/auth/cli/***
   npm error After authenticating, your token can be retrieved from: https://registry.npmjs.org/-/v1/done?authId=***
   ```

   处理方式二选一：
   - **按提示做一次**：浏览器打开那个授权链接完成验证，回到终端重跑 `npm publish`（会话式授权，最省事）；
   - **换个 token**：在 npmjs 账号里删掉现有 token，重新建一个 **Granular Access Token** 并**勾选
     `Bypass 2FA`**（或建一个 **Automation** 类型的 classic token，它天然绕过 2FA），写回
     `//registry.npmjs.org/:_authToken=<新 token>`，之后发布不再要 OTP。

   `npm publish --dry-run` **不需要** OTP，所以它全绿并不代表真发布能过 —— 这两步要分开看。

3. **`unpublish` / `deprecate` 这类破坏性写操作，任何 token 都走不通（2026-10-02 实测，三次）**：
   - Granular token **没勾** Bypass（`bypass_2fa: false`）→ `npm unpublish` 回 `EOTP`；
   - Granular token **勾了** Bypass（`bypass_2fa: true`）→ npm 直接**禁止**该操作：

     ```
     403 Granular access tokens that bypass two-factor authentication may not perform this action.
     https://gh.io/npm-gat-bypass2fa-deprecation
     ```

     即 Bypass 只对 `publish` 放行，**对 unpublish/deprecate 是明确禁止**；
   - **授权 URL 与发起那一次命令的进程绑定**：在非交互终端（脚本 / 后台任务）里跑，npm 不会挂起等待，
     会直接退出并打印 URL —— 那个 URL 对下一次命令无效（authId 变了）。
   - 因此这类操作**只有一条可行路径**：**在交互终端里亲自跑那一条命令、让它挂着**，浏览器打开它打印的
     `https://www.npmjs.com/auth/cli/…` 并确认，命令随后自行完成。
     （2026-10-02 就是这么把 `0.6.0` 撤掉的：`npm unpublish …@0.6.0` → 浏览器确认 → 完成后 registry 上只剩 `0.6.1`。）

**另一个坑**：
仓库里的 `publishConfig.registry` 会让 `npm publish` 自动改投 npmjs，所以**不要**在发布命令里再写
`--registry=https://registry.npmmirror.com/`；想显式一点就写 `--registry=https://registry.npmjs.org/`。
`get` 类命令（`npm view` / `npm whoami`）不受 `publishConfig` 影响，要自己带 `--registry`。

---

## 1. 发布前检查（全部必过）

工作区必须干净，只包含本次发布该有的改动（提交纪律：**只写明确路径，禁止 `git add -A`**）：

```powershell
Set-Location "D:\dsh\a"
git status --short
node -e "console.log(require('./package.json').version)"   # 必须等于本次要发的版本
```

静态门禁（**不需要 WPS**，CI 跑的是同一批）：

```powershell
cd mcp; npx tsc; cd ..                     # 构建产物要与源码一致
git diff --exit-code -- mcp/dist
powershell -NoProfile -File scripts\build-host-actions.ps1
git diff --exit-code -- host/wps-actions.ps1
node scripts\gen-tool-surface.mjs;     git diff --exit-code -- spec
node scripts\gen-skill-tools.mjs;      git diff --exit-code -- skills
node scripts\gen-tool-coverage.mjs;    git diff --exit-code -- docs/tool-coverage.md
npm run lint
node scripts\verify.mjs --static
node scripts\param-contract.mjs;       git diff --exit-code -- docs\param-contract.md
npm run verify:package                 # 打包产物运行期冒烟：装一遍 + 启动 MCP server 握手（必跑，见 §3.1；CI 也跑）
# 逐文件跑测试前先清掉插件自己发布的两个环境变量，否则 plugin.test.mjs 会假失败：
# $env:WPS_OFFICE_MCP_ENTRY=$null; $env:WPS_OFFICE_HOST_SCRIPT=$null
```

真机部分（**可选但建议**，需要在跑着 WPS 的机器上）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1   # 整轮 + 回收无头 WPS 孤儿
node scripts\e2e.mjs --setup --timeout 420 --profile wpse2e                  # 一键端到端
node scripts\accept-install.mjs                                              # 全新 profile 安装验收
```

## 2. 打包内容对账（发布前最后一关）

```powershell
npm pack --dry-run --json | ConvertFrom-Json | ForEach-Object {
  "files=" + $_.entryCount + "  packed=" + [math]::Round($_.size/1KB,1) + "KB  unpacked=" + [math]::Round($_.unpackedSize/1MB,2) + "MB"
}
```

**期望值**（**2026-10-03 当场实测**）：`files=290 / entryCount=290  packed≈580KB  unpacked≈3.04MB`。

> 这行数字**会随仓库增长漂移**（0.6.1 发布时是 288 + `entryCount` 289、3.03MB；现在源码与脚本变多，已是 290 / 3.04MB）。
> 所以它只是**量具的读数**，不是断言：真正的断言在 §3.1 的 `verify-package.mjs` 里（包内必需文件 + server 真的能起来）。
> 每次发布照上面那条命令**重新量一次**，顺手把这一行更新掉，别照抄本文的历史数字。

必须出现、且一个都不能少：

| 路径 | 为什么 |
| --- | --- |
| `plugin.js` | DSH 入口，负责把包内绝对路径写进环境变量 |
| `cordis.patch.yml` | 接线：`wps-office-next-plugin` + `mcp-wps-office-next` 两个 id |
| `mcp/dist/index.js` | MCP server 入口（预构建，安装即用） |
| `mcp/package.json` | **`mcp/` 的模块边界声明**。`mcp/dist` 是 CommonJS 产物，少了它就会继承根 `package.json` 的 `"type": "module"`，server 一启动就抛 `exports is not defined in ES module scope` —— 0.6.0 就是这么坏的 |
| `mcp/scripts/wps-com.ps1` | 桥的真源（排错时对照 action 契约） |
| `host/wps-com-host.ps1`、`host/wps-actions.ps1` | 常驻 COM 宿主与生成物；**必须带 UTF-8 BOM** |
| `skills/*/SKILL.md`（4 份）、`reference.md`（4 份） | 模型面向的技能文档 |
| `scripts/doctor.mjs`、`scripts/verify.mjs` | 用户侧自检与门禁 |

必须**没有**：

| 路径 | 为什么不能有 |
| --- | --- |
| `mcp/node_modules/**` | 开发依赖（jest / ts-node / typescript 等 329 个顶层目录）。0.5.4 时曾混进去：10,280 文件 / 76.64 MB |
| `mcp/src/**`、`test/**`、`docs/**`、`baseline/**`、`spec/**` | 开发资料，不进包（`spec/` 只是生成物，运行时不读） |

顺带核对两件事（在**解开的包**里看，不是看源码）：

```powershell
# 1) BOM 仍在（缺 BOM 会让 PS 5.1 把中文读成乱码、宿主直接解析失败）
foreach ($f in 'host\wps-com-host.ps1','host\wps-actions.ps1') {
  $b = [System.IO.File]::ReadAllBytes("$f")[0..2]
  "$f BOM=" + (($b[0] -eq 0xEF) -and ($b[1] -eq 0xBB) -and ($b[2] -eq 0xBF))
}
# 2) 桥仍是纯 CRLF 无 BOM
$b = [System.IO.File]::ReadAllBytes('mcp\scripts\wps-com.ps1')
"bridge first3=" + ($b[0..2] -join ',')
```

## 3. 真·收包验证（**v0.6.1 起改为强制跑脚本**）

### 3.1 一条命令：`node scripts\verify-package.mjs`（**发布前必跑**）

它自己完成 `npm pack` → 装进临时目录 → **真的把 MCP server 拉起来做一次 JSON-RPC 握手** → 断言 `tools/list`
广告 69 个工具、`wps_status` 在列。**不需要 WPS**，全程只碰临时目录。**11 个断言**全绿会打印
`PACKAGE RUNTIME OK`（2026-10-03 实测：12.4 秒、11/11）。

> **CI 也会跑它**（2026-10-03 起，D6-B）：`.github/workflows/ci.yml` 里新增 gate `npm run verify:package`——
> 所以这一关不再依赖「人记得跑」。本地仍建议在**改动 `package.json` 的 `files` 时**手动跑一次。
>
> **务必照上面这条命令（经 `npm run`）跑**，别改成 `node scripts\verify-package.mjs`：两者的环境不同，
> 这个脚本曾在「直接调用全绿、经 `npm run` 必挂」的差异下藏了一个真缺陷（FIXES 85 —— 现已修，但仍按 CI 的命令验）。

> **为什么加这一关（FIXES 84）**：0.6.0 发布时只核对了「文件在不在」，而漏发的 `mcp/package.json` 是个
> **运行期**缺件 —— 文件清单检查永远抓不到它，唯一能抓住的办法就是把 server 真的启动一次。

### 3.2 发布后 / 排错：`node scripts\probe-installed.mjs [插件根目录]`

对**任意一个已安装的副本**（默认取 `desktop` profile 里那份）做同样的握手，并**真的调一次 `wps_status`**
（需要 WPS）。它的价值是当场分辨「包坏了」还是「环境没起 WPS」：坏包会在
`HANDSHAKE FAILED: timeout waiting for initialize` 后面直接跟出 Node 的堆栈。

### 3.3 手工对照（可选）

把 tarball 装进一个**临时目录**（不要碰任何 profile），确认装出来的副本能自检：

```powershell
$tmp = 'D:\dsh\_pubcheck'; New-Item -ItemType Directory -Force -Path $tmp | Out-Null
Set-Location "D:\dsh\a"; npm pack --pack-destination $tmp
Set-Location $tmp
Set-Content package.json '{"name":"pubcheck","version":"1.0.0","private":true}' -Encoding utf8
npm install --no-audit --no-fund --ignore-scripts ./dsh-plugin-wps-office-next-0.6.0.tgz
node node_modules\dsh-plugin-wps-office-next\scripts\doctor.mjs     # 末尾必须是 DOCTOR OK
```

## 4. 发布

```powershell
Set-Location "D:\dsh\a"
npm publish --dry-run            # 再确认一次清单与体积；publishConfig 会把 registry 指到 npmjs
npm publish                      # 2FA 用户加 --otp <6 位码>；首次发布必要时加 --access public
```

> 仓库里**不放任何凭证**（D4-B）。`publishConfig.registry` 只记录目标源，不含认证信息。

## 5. 发布后验收（**v0.6.0 已按此执行，全过**）

```powershell
# 1) registry 上确实是发出去的那一份（integrity/shasum 要与第 4 步 dry-run 打印的一致）
npm view dsh-plugin-wps-office-next version dist-tags dist.integrity dist.fileCount --registry https://registry.npmjs.org --prefer-online

# 2) 全新隔离 profile 从 registry 真装一遍（不要用本机已有 profile 验）
dsh plugin --profile <临时名> add dsh-plugin-wps-office-next
dsh --profile <临时名> --dump-config | Select-String wps
# 装出来的副本里再跑一次自检：
Push-Location "$env:USERPROFILE\.dsh\profiles\<临时名>\node_modules\dsh-plugin-wps-office-next"
node scripts\doctor.mjs
Pop-Location
# 拆干净（必须带包名，不带会报 ERR_PNPM_MUST_REMOVE_SOMETHING）
dsh plugin --profile <临时名> remove dsh-plugin-wps-office-next
Remove-Item -Recurse -Force "$env:USERPROFILE\.dsh\profiles\<临时名>"
```

> 新包首次安装时 pnpm 会提示 `minimumReleaseAgeExclude` 并自动写一行到 `pnpm-workspace.yaml`，那是它的新版本
> 保护机制，**不是错误**。

然后才是真实用户路径。

```powershell
npm view dsh-plugin-wps-office-next version dist.tarball --registry https://registry.npmjs.org
```

在**一个全新的一次性 profile** 上按用户路径装一遍（这是验收分成败的唯一标准）：

```powershell
node scripts\accept-install.mjs --profile wpsnpm
```

然后才是真实用户路径（会动到你自己在用的 profile，**先问过用户再执行**）：

```powershell
dsh plugin --profile desktop add dsh-plugin-wps-office-next
dsh --profile desktop --dump-config | Select-String wps     # 应出现两个 id
```

**必须重启 DSH**（插件只在启动时加载），重启后新建会话调用 `wps_status`，确认 `connected: true`、
`advertisedTools: 69`、`registeredTools: 268`。

## 6. GitHub 侧（与 npm 同步）

```powershell
git add package.json package-lock.json mcp/package.json mcp/package-lock.json CHANGELOG.md README.md docs\HANDOFF.md docs\release-checklist.md
git commit -m "release: v0.6.0（上架 npm + 打包载荷 76.6MB → 3.0MB）"
git tag v0.6.0
git push origin main
git push origin v0.6.0
gh release create v0.6.0 --title "v0.6.0（上架 npm）" --notes "$(Select-String -Path CHANGELOG.md -Pattern '^## 0\.6\.0' -Context 0,30 | Out-String)"
```

README 里的 GitHub pin（`#v0.6.0`）指向的 tag 必须真实存在，否则回退路径是坏的。

## 7. 出问题怎么办

| 情况 | 处理 |
| --- | --- |
| `E402` / `E404` / `You do not have permission` | token 没配、没勾 Bypass 2FA、或包名被别人占了。先 `npm whoami --registry https://registry.npmjs.org` |
| `ENEEDAUTH` | 本机没有凭证 —— 见 §0 |
| 发错版本号（比如发了 0.6.1） | `npm publish` **不能覆盖同版本**；改号重发，或 `npm deprecate` 旧号 |
| 内容错了（多带文件 / 少带文件） | npm 允许 **72 小时内** `npm unpublish <pkg>@<version>`（且无人依赖）；超时或已被依赖就只能发新版本 —— 所以 §1–§3 必须先过 |
| 想撤回整个包 | `npm unpublish <pkg> --force`，但**别轻易用**：同版本号会被永久保留、无法再用 |

## 8. 每次发版的固定顺序

1. 改版本号（`package.json` + `mcp/package.json` + 两个 lockfile），写 CHANGELOG 与 README pin
2. §1 静态门禁全绿（+ 建议的真机）
3. §2 打包对账 → §3 收包验证
4. `npm publish`
5. §5 全新 profile 安装验收 → 真实 profile 安装 → 重启验证
6. §6 提交 / 打 tag / 推 / 发 Release
