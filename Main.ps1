# PowerShellスクリプト - GUIアプリケーション
# エンコーディング: UTF-8 BOM付
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 設定ファイルの読み込み
$script:configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $script:configPath) {
    $script:config = Get-Content $script:configPath -Encoding UTF8 | ConvertFrom-Json
}
else {
    Write-Host "設定ファイルが見つかりません: $script:configPath"
    exit 1
}

# グローバルログパスの設定
if ($script:config.GlobalLogPath) {
    if ([System.IO.Path]::IsPathRooted($script:config.GlobalLogPath)) {
        $script:globalLogPath = $script:config.GlobalLogPath
    }
    else {
        $script:globalLogPath = Join-Path $PSScriptRoot $script:config.GlobalLogPath
    }
}
else {
    # デフォルト設定（logsフォルダ）
    $script:globalLogPath = Join-Path $PSScriptRoot "logs"
}

# ログファイルのパス
$script:logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $script:logDir)) {
    New-Item -ItemType Directory -Path $script:logDir -Force | Out-Null
}

# セッションログファイルの初期化（起動ごとに単一のログファイルを生成）
$script:sessionStartTime = Get-Date -Format "yyyyMMdd_HHmmss"
$script:sessionLogFile = Join-Path $script:logDir "iko_gui_$script:sessionStartTime.log"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($script:sessionLogFile, "Session Started at $script:sessionStartTime`r`n", $utf8NoBom)

# グローバル変数
$script:currentPage = 0
$script:processesPerPage = 8
$script:processControls = @()
$script:processLogs = @{}
$script:pages = @()
$script:pageProcessCache = @()
$script:editMode = $false

# ページ設定の読み込み
if ($script:config.Pages) {
    $script:pages = $script:config.Pages
}
else {
    # 後方互換性のため、旧形式の設定もサポート
    if ($script:config.Processes) {
        $script:pages = @(@{
                Title     = if ($script:config.Title) { $script:config.Title } else { "" }
                JsonPath  = $null
                Processes = $script:config.Processes
            })
    }
    else {
        Write-Host "設定ファイルの形式が正しくありません"
        exit 1
    }
}

# 関数定義モジュールの読み込み
. $PSScriptRoot\Functions.ps1

# UIレイアウトモジュールの読み込み（アプリケーション起動）
. $PSScriptRoot\UILayout.ps1
