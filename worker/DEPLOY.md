# 部署说明：birthday-notice (Cloudflare Workers)

本程序由本地 `xlsx_to_json.py` 生成的 `birthday.json` 提供数据（**不依赖、不上传 xlsx**）。
Worker 每天定时检查阳历/阴历生日，提前 0~3 天通过 Bark 推送通知。

## 目录结构
```
birthday_notice/
├─ birthday.xlsx            # 本地数据源（不上传）
├─ xlsx_to_json.py          # 本地：xlsx -> json
└─ worker/
   ├─ wrangler.toml         # Worker 配置 + cron
   ├─ package.json
   └─ src/
      ├─ index.js           # Worker 主程序（只解析 json）
      └─ birthday.json      # 由 xlsx_to_json.py 生成
```

## 1. 本地更新数据（改 xlsx 后）
数据已放在 Cloudflare KV（命名空间 `BIRTHDAY`），Worker 运行时解析 KV 中的 json，**改数据无需重新部署**：

### 一键脚本（推荐，Windows）
`worker/update.ps1` 把「生成 json → 推 KV」合成一步，并从 Windows 环境变量读取 `CLOUDFLARE_API_TOKEN`（可选 `CLOUDFLARE_ACCOUNT_ID`）。Worker 的 `fetch` 每次拉取都按 KV 当前 `birthday.json` 现算，因此**只推 `birthday.json` 即可**，无需再处理 `calendar.ics`。
```powershell
cd worker
.\update.ps1            # 完整：xlsx->json + 推 KV(birthday.json)
.\update.ps1 -SkipJson  # 只推 KV（已直接改好 birthday.json 时，避免被 xlsx 覆盖回去）
# 若 PowerShell 禁止脚本：
powershell -ExecutionPolicy Bypass -File .\update.ps1
```
> `CLOUDFLARE_API_TOKEN` 保存在哪：见 `update.ps1` 底部说明（系统/用户环境变量，或 `setx CLOUDFLARE_API_TOKEN "..."`；Cloudflare 后台 My Profile → API Tokens 创建，权限勾 Workers Scripts(Edit)+Workers KV Storage(Edit)）。token 无 memberships 权限时另设 `CLOUDFLARE_ACCOUNT_ID`。

### 手工分步（任意系统）
```bash
python xlsx_to_json.py                                # 重新生成 src/birthday.json
# ⚠️ wrangler 4 起 kv 命令默认目标是【本地】KV，必须加 --remote 才写线上远程 KV，否则部署的 Worker 读不到
wrangler kv key put --binding=BIRTHDAY --remote birthday.json --path ./src/birthday.json   # 推送到远程 KV，立即生效（fetch 每次现算，无需删 calendar.ics）
```
> **自检是否真写到了远程**——⚠️ PowerShell 5.1 中文 locale 下，`Select-String`/`findstr` 对 UTF-8 管道输出常因被当成 GBK 解码而乱码、导致 `王创` 漏匹配（**空 ≠ KV 没数据**）。最可靠的坐实法是把远程 KV 落盘、按字节查，绕过管道编码：
> ```powershell
> node node_modules\wrangler\bin\wrangler.js kv key get birthday.json --binding=BIRTHDAY --remote > C:\Users\zou\kv_check.json
> # 然后用 notepad 打开 kv_check.json 搜“王创”，或让 WorkBuddy 直接 Read 该文件确认
> ```
> （不带 `--remote` 读的是本地 KV，与线上 Worker 无关，不要用来判断线上数据。）
> 若未启用 KV（注释掉 wrangler.toml 里的 [[kv_namespaces]]），则改为 `wrangler deploy` 重新发布（数据随包内置）。

## 2. 安装依赖 & 部署
```bash
cd worker
npm install
wrangler login                 # 浏览器登录 Cloudflare（首次需要）
wrangler secret put BARK_KEY   # 输入你的 Bark key（即 api.day.app/<KEY> 中的 KEY）
wrangler secret put CAL_TOKEN  # 输入日历订阅令牌（任意随机串，如 `openssl rand -hex 12`）
wrangler deploy
```

## 3. 关键配置
- **BARK_KEY**：密钥，通过 `wrangler secret put BARK_KEY` 设置，不要写进代码/仓库。
- **BARK_BASE**：普通变量，默认 `https://api.day.app/`，在 `wrangler.toml` 的 `[vars]` 中可改。
- **定时触发**：`[triggers] crons = ["0 16 * * *"]`，即每天 **中国 00:00**（Cloudflare cron 用 UTC，中国 = UTC+8，故 16:00 UTC）。
  - 改时间在该行调整，如每天 08:00 中国 = `0 0 * * *` 改为 `0 0`? 中国 08:00 = UTC 00:00 → `"0 0 * * *"`。
- **手动触发 / 自检**：访问 `https://<你的子域>.workers.dev` 即运行一次检查，返回 JSON 结果（便于测试，不会因 cron 等待）。

## 4. 逻辑说明（对照原 py）
- 阳历生日：今年对应阳历日期，距今天 0~3 天则提醒。
- 阴历生日：用 `lunar-javascript` 把今年农历 月/日 转为阳历（对照原 `zhdate`）。**与原 py 严格一致**：若某农历年/月实际没有该日（如 农历七月只有29天却写了30日），与原 `zhdate` 抛错→整行 skip 的行为一致，**跳过该年**（不进位到下月、不产生错误日期）。这是与原 py 对齐的关键点。
- 已过的今年生日不再提醒，明年自然重新进入窗口。
- 通知文案与原脚本一致：当天 `「张三今天过44岁生日,阴历...」`，非当天 `「张三2026-09-22(3天后)过44岁生日,阴历...」`。

## 5. 日历订阅（ICS，滚动窗口 + 预生成）
Worker 提供 iCalendar 订阅源，可添加到手机/电脑日历 App。设计为「**每次拉取都按 KV 当前数据现算**」，不依赖可能陈旧的预生成文件。

- **订阅地址**：`https://<子域>.workers.dev/?token=<CAL_TOKEN>`
  - 例：`https://birthday-notice.sailing-sailor.workers.dev/?token=8a7f127dd388c070ea61a830`
- **token 校验**：`token` 缺失或错误一律返回 **404**；只有与 `CAL_TOKEN` 一致才返回 `text/calendar`。
- **工作原理**：
  1. 每天定时（cron = 中国 00:00）执行 `scheduled`：用 KV 中当前 `birthday.json` 重新生成「前后 60 天」日历，写入 KV 的 `calendar.ics`（含 metadata：生成时间 / 中国日期 / 窗口天数）作为兜底缓存。
  2. 客户端拉取订阅地址时，`fetch` **每次都按 KV 中“当前”的 `birthday.json` 现算并返回**（`Cache-Control: max-age=60`），不再信任那份 `calendar.ics` 缓存文件——改了数据后只要 KV 里的 `birthday.json` 是最新的，订阅下次刷新（≤60s）即见，避免“改了数据却没刷新”的失效模式。
- **内容（滚动 ±60 天窗口，共 121 天）**：
  - 阳历生日：取当前年/前一年/后一年的具体月日，落在窗口内的生成单条 `VEVENT`（无 RRULE 重复，因为窗口每天滚动刷新）。
  - 阴历生日：用 `lunar-javascript` 把农历 月/日 转阳历（对照原 `zhdate`）；该农历月没有这一天（如七月只有29天却写30日）按年跳过，与原 py 严格一致。同样只保留落在窗口内的年份。
  - 没有生日日期的空行 / 未知类型会被跳过。
  - 每个 `VEVENT` 带 `SUMMARY`（姓名 + 生日 + 年龄）和 `DESCRIPTION`（阳历/农历原始日期 + 备注）。
- **窗口可调**：改 `src/index.js` 顶部 `CAL_WINDOW`（默认 60）即可改前后天数。
- **数据来源**：生日数据读 KV（`BIRTHDAY` 里的 `birthday.json`）；日历文件存同一命名空间的 `calendar.ics`。
- 改 xlsx 后：`python xlsx_to_json.py` → `wrangler kv key put --binding=BIRTHDAY --remote birthday.json "$(cat src/birthday.json)"`，**次日 00:00 自动重新生成日历**生效。若想立刻生效，也可手动 `wrangler kv key put --binding=BIRTHDAY --remote calendar.ics "$(本地生成的ics)"` 或等待首次拉取兜底生成。
  - ⚠️ **wrangler 4 必加 `--remote`**：kv 命令默认写「本地」KV，不加则部署的 Worker（读远程）拿不到，会退回打包内置旧数据。
- ⚠️ Bark 每日提前推送目前仍保留（`scheduled` 里 `checkBirthdays` 调用）；若只想保留日历，删掉该调用即可。

## 6. 本地调试（不部署）
```bash
wrangler dev        # 本地起服务，访问 http://127.0.0.1:8787 触发检查
```

## 7. 自动部署（GitHub Actions）
仓库已包含 `.github/workflows/deploy.yml`：push 到 `main` 且 `worker/**` 或 workflow 自身变动时，自动 `npm ci` + `npx wrangler deploy`（在 `worker/` 目录内执行）。

### 首次只需做一次：配置仓库 Secrets
仓库页面 → **Settings → Secrets and variables → Actions → New repository secret**，添加：
- **`CLOUDFLARE_API_TOKEN`**：Cloudflare API Token（权限需 `Workers Scripts(Edit)` + `Workers KV Storage(Edit)`）。
  取你本地已有的同一个 token：Cloudflare 后台 `My Profile → API Tokens`；或本机 Windows 环境变量 `CLOUDFLARE_API_TOKEN` 的值。
- **`CLOUDFLARE_ACCOUNT_ID`**（可选但建议）：账号 ID `6b26e121057fd094c5e176f5070b2338`。当 token 无 memberships 权限时必须提供，否则 deploy 报 `/memberships` 错误。

### 之后
- 改 `worker/` 代码 → push 到 `main` → 自动部署，几秒到一两分钟完成。
- 也可在 **Actions** 页面手动 **Run workflow** 立即触发（workflow 已开启 `workflow_dispatch`）。

### 注意
- 首次 push 时 secret 尚未配置，workflow 会失败；配好 secret 后重跑（或下次 push）即成功。
- Worker 上的密钥 `BARK_KEY`、`CAL_TOKEN` 已通过 `wrangler secret put` 存于 Cloudflare，**自动部署只更新代码，不会清除这些 secret**。
- 数据更新走 `update.ps1` → KV，不触发部署；只有 `worker/` 代码改动才触发部署。

## 8. Cloudflare API Token（权限与存储位置）

`update.ps1` / `wrangler deploy` 从 **Windows 环境变量**读取 `CLOUDFLARE_API_TOKEN`，绝不写进代码或仓库。

### 创建 Token（Cloudflare 后台）
- 入口：`My Profile → API Tokens → Create Token → Create Custom Token`
- 账号范围：选你自己的账号（`6b26e121057fd094c5e176f5070b2338`）
- **权限**（Account 级别，最小集）：

  | 权限 | 操作 | 用途 |
  |---|---|---|
  | Cloudflare Workers Scripts | Edit | `wrangler deploy` 部署 Worker |
  | Workers KV Storage | Edit | `wrangler kv key put` 推 `birthday.json` |

- （可选）再加 **Account Settings → Read**：避免 wrangler 在无 `CLOUDFLARE_ACCOUNT_ID` 时调 `/memberships` 报错。更省事是直接把账号 ID 也设成环境变量（见下）。
- TTL：建议设长（1 年或不过期）；创建后**立即复制**，页面关闭不可再见。

### 存储位置（两处，互相独立）
1. **本机 Windows 环境变量**（给 `update.ps1` / 本地 `wrangler deploy`）：
   ```powershell
   setx CLOUDFLARE_API_TOKEN "你的token值"                                    # 用户环境变量，重开终端生效
   setx CLOUDFLARE_ACCOUNT_ID "6b26e121057fd094c5e176f5070b2338"             # 建议一并设置，跳过 /memberships 查询
   # 图形界面：设置 → 系统 → 关于 → 高级系统设置 → 环境变量 → 用户变量 → 新建
   # 仅本次终端临时用： $env:CLOUDFLARE_API_TOKEN = "你的token值"
   ```
2. **GitHub 仓库 Secret**（给 Actions 自动部署）：仓库 `Settings → Secrets and variables → Actions → New repository secret`，Name=`CLOUDFLARE_API_TOKEN`，值同上；`CLOUDFLARE_ACCOUNT_ID` 已硬编码在 `deploy.yml`，无需再配。
