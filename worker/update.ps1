<#
.SYNOPSIS
  一键更新生日日历到 Cloudflare：生成 json -> 推 KV -> 删旧日历。
.DESCRIPTION
  1) 可选：运行 xlsx_to_json.py 把 birthday.xlsx 转成 birthday.json
  2) 把 birthday.json 推到 KV（命名空间 BIRTHDAY）
  3) 删除旧 calendar.ics，使下次订阅拉取自动重建（或等次日 00:00 定时重建）
  凭据从 Windows 环境变量读取 CLOUDFLARE_API_TOKEN（可选 CLOUDFLARE_ACCOUNT_ID）。
.PARAMETER SkipJson
  跳过 xlsx->json 步骤（用于你已经直接改好 birthday.json 的情况，避免被 xlsx 覆盖回去）。
.EXAMPLE
  .\update.ps1            # 完整流程
  .\update.ps1 -SkipJson  # 只推 KV + 删旧日历
#>
param(
  [switch]$SkipJson
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$projDir   = Split-Path $scriptDir -Parent

# ---------- 从 Windows 环境变量读取 Cloudflare 凭据 ----------
function Get-WinEnv($name) {
  $v = [Environment]::GetEnvironmentVariable($name, "User")
  if ($v) { return $v }
  $v = [Environment]::GetEnvironmentVariable($name, "Machine")
  if ($v) { return $v }
  if (Test-Path "env:$name") { return (Get-Item "env:$name").Value }
  return $null
}

$token = Get-WinEnv "CLOUDFLARE_API_TOKEN"
if (-not $token) {
  Write-Host ""
  Write-Host "未找到 CLOUDFLARE_API_TOKEN。" -ForegroundColor Red
  Write-Host "请先在 Windows 环境变量里设置它（详见本脚本底部“凭据保存位置”说明），" -ForegroundColor Yellow
  Write-Host "或本次临时设置：`$env:CLOUDFLARE_API_TOKEN = '你的token值'`，再运行本脚本。" -ForegroundColor Yellow
  exit 1
}
$env:CLOUDFLARE_API_TOKEN = $token   # 传给 wrangler

# 账号 ID：非必须；若你的 token 没有 memberships 权限则必须提供，否则 kv/deploy 会报 /memberships 错误
$acct = Get-WinEnv "CLOUDFLARE_ACCOUNT_ID"
if ($acct) { $env:CLOUDFLARE_ACCOUNT_ID = $acct }

# wrangler 本地路径（不依赖 PATH）
$wrangler = Join-Path $scriptDir "node_modules/wrangler/bin/wrangler.js"
if (-not (Test-Path $wrangler)) {
  Write-Error "未找到 $wrangler，请先在 worker 目录执行 npm install"
  exit 1
}

# ---------- 1) 生成 json（可选） ----------
if (-not $SkipJson) {
  Write-Host "==> 生成 birthday.json (xlsx -> json)" -ForegroundColor Cyan
  & python (Join-Path $projDir "xlsx_to_json.py")
  if ($LASTEXITCODE -ne 0) { Write-Error "xlsx_to_json.py 执行失败（请确认 python 与 openpyxl 可用）"; exit 1 }
} else {
  Write-Host "==> 跳过 xlsx->json（使用现有 birthday.json）" -ForegroundColor Cyan
}

# ---------- 2) + 3) 推 KV 并删旧日历（在 worker 目录执行，--path 相对 worker/） ----------
Push-Location $scriptDir
try {
  Write-Host "==> 上传 birthday.json 到 KV (BIRTHDAY)" -ForegroundColor Cyan
  & node $wrangler kv key put birthday.json --binding=BIRTHDAY --path ./src/birthday.json
  if ($LASTEXITCODE -ne 0) { Write-Error "上传 birthday.json 失败"; exit 1 }

  Write-Host "==> 删除旧 calendar.ics（下次订阅拉取自动重建，或次日 00:00 定时重建）" -ForegroundColor Cyan
  & node $wrangler kv key delete calendar.ics --binding=BIRTHDAY
  if ($LASTEXITCODE -ne 0) { Write-Warning "删除 calendar.ics 失败（可能本就不存在，可忽略）" }
} finally {
  Pop-Location
}

Write-Host ""
Write-Host "完成。订阅地址（token 校验失败返回 404）：" -ForegroundColor Green
Write-Host "  https://birthday-notice.sailing-sailor.workers.dev/?token=8a7f127dd388c070ea61a830" -ForegroundColor Green

<#
================================================================
凭据保存位置（CLOUDFLARE_API_TOKEN 在哪里设置）
----------------------------------------------------------------
Cloudflare API Token 是“永久保存在你 Windows 系统里的环境变量”，
脚本运行时从 Windows 环境变量读取，不会写进代码或仓库。

【方式一：图形界面（推荐，长期有效）】
  设置 -> 系统 -> 关于 -> 高级系统设置 -> 环境变量
  -> 在“用户变量”或“系统变量”里 新建：
       变量名： CLOUDFLARE_API_TOKEN
       变量值： 你的 token 值（在 Cloudflare 后台复制）
  -> 确定后“重开” PowerShell 终端即生效。

【方式二：PowerShell 命令（写入用户环境变量）】
  setx CLOUDFLARE_API_TOKEN "你的token值"
  （setx 写入“用户”环境变量，重开终端后生效）

【方式三：仅本次终端临时用（关掉终端失效）】
  $env:CLOUDFLARE_API_TOKEN = "你的token值"

【获取 token】
  Cloudflare 后台 -> My Profile -> API Tokens -> Create Token
  权限建议勾选：
    - Account > Workers Scripts          (Edit)
    - Account > Workers KV Storage       (Edit)
  创建后复制得到的字符串即 CLOUDFLARE_API_TOKEN。

【CLOUDFLARE_ACCOUNT_ID（可选）】
  如果你的 token 没有 memberships 权限，kv/deploy 会报错，
  这时同样建一个 Windows 环境变量：
       变量名： CLOUDFLARE_ACCOUNT_ID
       变量值： 你的 Cloudflare 账号 ID
                （后台右下角 / 任意资源 URL 里的 ``:/account/xxxxx``）

【运行本脚本】
  PowerShell 首次可能禁止脚本，用下面方式运行：
    powershell -ExecutionPolicy Bypass -File .\update.ps1
  或先设一次：
    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
================================================================
#>
