<#
.SYNOPSIS
Creates a GitHub Release and uploads dist\TradeXmlStudio.exe.

.EXAMPLE
.\publish-release.ps1 1.0.1

.EXAMPLE
.\publish-release.ps1 v1.0.1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Version
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repository = 'ABCsfree/single_window'
$targetBranch = 'main'
$executablePath = Join-Path $PSScriptRoot 'dist\TradeXmlStudio.exe'

$normalizedVersion = $Version.Trim()
if ($normalizedVersion.StartsWith('v', [System.StringComparison]::OrdinalIgnoreCase)) {
    $normalizedVersion = $normalizedVersion.Substring(1)
}

if ($normalizedVersion -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$') {
    throw "版本号格式无效：$Version。请使用 1.2.3 或 1.2.3-beta.1。"
}

$tag = "v$normalizedVersion"

if (-not (Test-Path -LiteralPath $executablePath -PathType Leaf)) {
    throw "找不到待上传文件：$executablePath。请先生成或复制 TradeXmlStudio.exe 到 dist 目录。"
}

$executable = Get-Item -LiteralPath $executablePath
if ($executable.Length -eq 0) {
    throw "待上传文件为空：$executablePath"
}

$sizeMiB = [Math]::Round($executable.Length / 1MB, 2)
$asset = "$($executable.FullName)#TradeXmlStudio.exe (Windows x64)"
$title = "TradeXmlStudio $tag"

Write-Host "准备发布 $title"
Write-Host "文件：$($executable.FullName) ($sizeMiB MiB)"

if (-not $PSCmdlet.ShouldProcess("$repository / $tag", '创建 GitHub Release 并上传 EXE')) {
    return
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw '未找到 GitHub CLI（gh）。请先安装并执行 gh auth login。'
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw '未找到 Git。'
}

& gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
    throw 'GitHub CLI 尚未登录。请先执行 gh auth login。'
}

& gh repo view $repository --json nameWithOwner *> $null
if ($LASTEXITCODE -ne 0) {
    throw "无法访问 GitHub 仓库 $repository，请检查网络、登录账号和仓库权限。"
}

$workingTreeChanges = & git -C $PSScriptRoot status --porcelain
if ($LASTEXITCODE -ne 0) {
    throw "无法读取 Git 仓库状态：$PSScriptRoot"
}

if ($workingTreeChanges) {
    throw '仓库中存在尚未提交的文件。请先提交并推送源码，再创建 Release。'
}

$localHead = (& git -C $PSScriptRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) {
    throw '无法读取本地 HEAD。'
}

$remoteHeadLine = & git -C $PSScriptRoot ls-remote origin "refs/heads/$targetBranch"
if ($LASTEXITCODE -ne 0 -or -not $remoteHeadLine) {
    throw "无法读取远端 origin/$targetBranch。请确认远端存在并可访问。"
}

$remoteHead = ($remoteHeadLine -split '\s+')[0]
if ($localHead -ne $remoteHead) {
    throw "本地 HEAD 与 origin/$targetBranch 不一致。请切换到 $targetBranch 并完成 git push 后重试。"
}

& gh release view $tag --repo $repository *> $null
if ($LASTEXITCODE -eq 0) {
    throw "Release $tag 已存在。请使用新的版本号。"
}

$releaseArguments = @(
    'release'
    'create'
    $tag
    $asset
    '--repo'
    $repository
    '--target'
    $targetBranch
    '--title'
    $title
    '--generate-notes'
    '--latest'
)

& gh @releaseArguments
if ($LASTEXITCODE -ne 0) {
    throw "创建 Release $tag 或上传 EXE 失败。"
}

Write-Host "发布完成：https://github.com/$repository/releases/tag/$tag" -ForegroundColor Green
Write-Host "最新版下载：https://github.com/$repository/releases/latest/download/TradeXmlStudio.exe" -ForegroundColor Green
