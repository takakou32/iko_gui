# PowerShellスクリプト - GUIアプリケーション
# エンコーディング: Shift-JIS

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 設定ファイルの読み込み
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json
} else {
    Write-Host "設定ファイルが見つかりません: $configPath"
    exit 1
}

# ログファイルのパス
$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# グローバル変数
$script:currentPage = 0
$script:processesPerPage = 8
$script:processControls = @()
$script:processLogs = @{}
$script:pages = @()
$script:pageProcessCache = @{}

# ページ設定の読み込み
if ($config.Pages) {
    $script:pages = $config.Pages
} else {
    # 後方互換性のため、旧形式の設定もサポート
    if ($config.Processes) {
        $script:pages = @(@{
            Title = if ($config.Title) { $config.Title } else { "" }
            JsonPath = $null
            Processes = $config.Processes
        })
    } else {
        Write-Host "設定ファイルの形式が正しくありません"
        exit 1
    }
}

# 現在のページのプロセス一覧を取得
function Get-CurrentPageProcesses {
    if ($script:currentPage -ge $script:pages.Count) {
        return @()
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # JsonPathが指定されている場合は、そのJSONファイルを読み込む
    if ($pageConfig.JsonPath) {
        $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        } else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
        
        if (Test-Path $jsonPath) {
            try {
                $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                if ($pageJson.Processes) {
                    return $pageJson.Processes
                } else {
                    Write-Log "JSONファイルにProcessesが含まれていません: $jsonPath" "WARN"
                    return @()
                }
            } catch {
                Write-Log "JSONファイルの読み込みに失敗しました: $jsonPath - $($_.Exception.Message)" "ERROR"
                return @()
            }
        } else {
            Write-Log "JSONファイルが見つかりません: $jsonPath" "ERROR"
            return @()
        }
    }
    
    # JsonPathが指定されていない場合は、直接Processesを使用（後方互換性）
    if ($pageConfig.Processes) {
        return $pageConfig.Processes
    }
    
    return @()
}

# ログ出力関数
function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [int]$ProcessIndex = -1)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # プロセス固有のログファイル
    if ($ProcessIndex -ge 0) {
        $processLogFile = Join-Path $logDir "process_${script:currentPage}_${ProcessIndex}.log"
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::AppendAllText($processLogFile, $logMessage + "`r`n", $utf8NoBom)
        $script:processLogs["${script:currentPage}_${ProcessIndex}"] = $processLogFile
    }
    
    # GUIのログ表示エリアに追加
    $script:logTextBox.AppendText("$logMessage`r`n")
    $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length
    $script:logTextBox.ScrollToCaret()
    
    Write-Host $logMessage
}

# Batファイル実行関数
function Invoke-BatchFile {
    param(
        [string]$BatchPath,
        [string]$DisplayName,
        [int]$ProcessIndex
    )
    
    if (-not (Test-Path $BatchPath)) {
        Write-Log "バッチファイルが見つかりません: $BatchPath" "ERROR" $ProcessIndex
        return $false
    }
    
    Write-Log "バッチファイルを実行中: $DisplayName ($BatchPath)" "INFO" $ProcessIndex
    
    try {
        $stdoutFile = Join-Path $logDir "process_${script:currentPage}_${ProcessIndex}_stdout.log"
        $stderrFile = Join-Path $logDir "process_${script:currentPage}_${ProcessIndex}_stderr.log"
        
        $process = Start-Process -FilePath $BatchPath -WorkingDirectory (Split-Path $BatchPath) -Wait -NoNewWindow -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
        
        if ($process.ExitCode -eq 0) {
            Write-Log "バッチファイルの実行が完了しました: $DisplayName (終了コード: $($process.ExitCode))" "INFO" $ProcessIndex
            return $true
        } else {
            Write-Log "バッチファイルの実行でエラーが発生しました: $DisplayName (終了コード: $($process.ExitCode))" "ERROR" $ProcessIndex
            return $false
        }
    } catch {
        Write-Log "バッチファイルの実行中に例外が発生しました: $DisplayName - $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# CSVファイル移動関数
function Move-CsvFiles {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$ProcessIndex
    )
    
    if (-not (Test-Path $SourcePath)) {
        Write-Log "ソースパスが見つかりません: $SourcePath" "ERROR" $ProcessIndex
        return $false
    }
    
    if (-not (Test-Path $DestinationPath)) {
        Write-Log "移動先ディレクトリを作成します: $DestinationPath" "INFO" $ProcessIndex
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
    
    try {
        $csvFiles = Get-ChildItem -Path $SourcePath -Filter "*.csv" -File
        
        if ($csvFiles.Count -eq 0) {
            Write-Log "CSVファイルが見つかりません: $SourcePath" "WARN" $ProcessIndex
            return $false
        }
        
        foreach ($file in $csvFiles) {
            $destFile = Join-Path $DestinationPath $file.Name
            Move-Item -Path $file.FullName -Destination $destFile -Force
            Write-Log "CSVファイルを移動しました: $($file.Name) -> $DestinationPath" "INFO" $ProcessIndex
        }
        
        Write-Log "CSVファイルの移動が完了しました (移動数: $($csvFiles.Count))" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "CSVファイルの移動中にエラーが発生しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# プロセス実行関数
function Start-ProcessFlow {
    param([int]$ProcessIndex)
    
    $currentProcesses = Get-CurrentPageProcesses
    $processConfig = $currentProcesses[$ProcessIndex]
    if (-not $processConfig) {
        return
    }
    
    $executeButton = $script:processControls[$ProcessIndex].ExecuteButton
    $executeButton.Enabled = $false
    $script:logTextBox.Clear()
    
    Write-Log "プロセスを開始します: $($processConfig.Name)" "INFO" $ProcessIndex
    
    $allSuccess = $true
    
    # バッチファイルの実行
    if ($processConfig.BatchFiles) {
        foreach ($batch in $processConfig.BatchFiles) {
            $batchPath = if ([System.IO.Path]::IsPathRooted($batch.Path)) {
                $batch.Path
            } else {
                Join-Path $PSScriptRoot $batch.Path
            }
            
            $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $ProcessIndex
            if (-not $result) {
                $allSuccess = $false
            }
            
            # 実行間隔（設定されている場合）
            if ($processConfig.ExecutionDelay -and $processConfig.ExecutionDelay -gt 0) {
                Start-Sleep -Seconds $processConfig.ExecutionDelay
            }
        }
    }
    
    # CSVファイルの移動
    if ($processConfig.CsvMoveOperations) {
        foreach ($csvOp in $processConfig.CsvMoveOperations) {
            $sourcePath = if ([System.IO.Path]::IsPathRooted($csvOp.Source)) {
                $csvOp.Source
            } else {
                Join-Path $PSScriptRoot $csvOp.Source
            }
            
            $destPath = if ([System.IO.Path]::IsPathRooted($csvOp.Destination)) {
                $csvOp.Destination
            } else {
                Join-Path $PSScriptRoot $csvOp.Destination
            }
            
            $result = Move-CsvFiles -SourcePath $sourcePath -DestinationPath $destPath -ProcessIndex $ProcessIndex
            if (-not $result) {
                $allSuccess = $false
            }
        }
    }
    
    if ($allSuccess) {
        Write-Log "プロセスが正常に完了しました: $($processConfig.Name)" "INFO" $ProcessIndex
    } else {
        Write-Log "プロセスでエラーが発生しました: $($processConfig.Name)" "ERROR" $ProcessIndex
    }
    
    $executeButton.Enabled = $true
}

# ログ確認関数
function Show-ProcessLog {
    param([int]$ProcessIndex)
    
    $logKey = "${script:currentPage}_${ProcessIndex}"
    if ($script:processLogs.ContainsKey($logKey) -and (Test-Path $script:processLogs[$logKey])) {
        Start-Process notepad.exe -ArgumentList $script:processLogs[$logKey]
    } else {
        $processLogFile = Join-Path $logDir "process_${script:currentPage}_${ProcessIndex}.log"
        if (Test-Path $processLogFile) {
            Start-Process notepad.exe -ArgumentList $processLogFile
        } else {
            [System.Windows.Forms.MessageBox]::Show("ログファイルが見つかりません。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
}

# プロセスコントロールの更新
function Update-ProcessControls {
    # ページ遷移時にJSONファイルを読み込む
    $currentProcesses = Get-CurrentPageProcesses
    $totalPages = $script:pages.Count
    
    Write-Log "ページ $($script:currentPage + 1) のプロセスを読み込みました (プロセス数: $($currentProcesses.Count))" "INFO"
    
    # 既存のコントロールをクリア
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup) {
            $script:processPanel.Controls.Remove($ctrlGroup.NameTextBox)
            $script:processPanel.Controls.Remove($ctrlGroup.ExecuteButton)
            $script:processPanel.Controls.Remove($ctrlGroup.LogButton)
        }
    }
    $script:processControls = @()
    
    # 新しいコントロールを作成
    for ($i = 0; $i -lt $script:processesPerPage; $i++) {
        if ($i -lt $currentProcesses.Count) {
            $processConfig = $currentProcesses[$i]
            $row = [Math]::Floor($i / 2)
            $col = $i % 2
            
            # コントロールの位置計算
            $x = [int](10 + $col * 390)
            $y = [int](10 + $row * 60)
            
            # テキストボックス（タスク名表示用）
            $nameTextBox = New-Object System.Windows.Forms.TextBox
            $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
            $nameTextBox.Size = New-Object System.Drawing.Size(200, 40)
            $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
            $nameTextBox.ReadOnly = $true
            $nameTextBox.BackColor = [System.Drawing.Color]::White
            $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
            $nameTextBox.Multiline = $true
            $nameTextBox.Height = 40
            $script:processPanel.Controls.Add($nameTextBox)
            
            # 実行ボタン（オレンジ）
            $executeButton = New-Object System.Windows.Forms.Button
            $executeX = [int]($x + 210)
            $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
            $executeButton.Size = New-Object System.Drawing.Size(80, 40)
            $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
            $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
            $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
            $executeButton.FlatAppearance.BorderSize = 1
            $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
            $processIdx = $i
            $executeButton.Add_Click({
                Start-ProcessFlow -ProcessIndex $processIdx
            })
            $script:processPanel.Controls.Add($executeButton)
            
            # ログ確認ボタン（緑）
            $logButton = New-Object System.Windows.Forms.Button
            $logX = [int]($x + 300)
            $logButton.Location = New-Object System.Drawing.Point($logX, $y)
            $logButton.Size = New-Object System.Drawing.Size(80, 40)
            $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
            $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
            $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
            $logButton.FlatAppearance.BorderSize = 1
            $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
            $logButton.Add_Click({
                Show-ProcessLog -ProcessIndex $processIdx
            })
            $script:processPanel.Controls.Add($logButton)
            
            $script:processControls += @{
                NameTextBox = $nameTextBox
                ExecuteButton = $executeButton
                LogButton = $logButton
            }
        }
    }
    
    # ページ情報の更新
    $script:pageLabel.Text = "ページ $($script:currentPage + 1) / $totalPages"
    
    # タイトルの更新
    $pageTitle = if ($script:pages[$script:currentPage].Title) { $script:pages[$script:currentPage].Title } else { if ($config.Title) { $config.Title } else { "1.V1 移行ツール適用" } }
    $script:titleLabel.Text = $pageTitle
}

# GUIフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "プロセス実行GUI"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

# ヘッダー部分（水色背景）
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(800, 50)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
$form.Controls.Add($headerPanel)

# タイトルラベル
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
$titleLabel.Text = if ($script:pages.Count -gt 0 -and $script:pages[0].Title) { $script:pages[0].Title } else { if ($config.Title) { $config.Title } else { "1.V1 移行ツール適用" } }
$titleLabel.Font = New-Object System.Drawing.Font("メイリオ", 12, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($titleLabel)
$script:titleLabel = $titleLabel

# 左矢印ボタン
$leftArrowButton = New-Object System.Windows.Forms.Button
$leftArrowButton.Location = New-Object System.Drawing.Point(700, 10)
$leftArrowButton.Size = New-Object System.Drawing.Size(40, 30)
$leftArrowButton.Text = "<"
$leftArrowButton.BackColor = [System.Drawing.Color]::Black
$leftArrowButton.ForeColor = [System.Drawing.Color]::White
$leftArrowButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$leftArrowButton.Font = New-Object System.Drawing.Font("メイリオ", 12, [System.Drawing.FontStyle]::Bold)
$leftArrowButton.Add_Click({
    if ($script:currentPage -gt 0) {
        $script:currentPage--
        Update-ProcessControls
    }
})
$headerPanel.Controls.Add($leftArrowButton)

# 右矢印ボタン
$rightArrowButton = New-Object System.Windows.Forms.Button
$rightArrowButton.Location = New-Object System.Drawing.Point(750, 10)
$rightArrowButton.Size = New-Object System.Drawing.Size(40, 30)
$rightArrowButton.Text = ">"
$rightArrowButton.BackColor = [System.Drawing.Color]::Black
$rightArrowButton.ForeColor = [System.Drawing.Color]::White
$rightArrowButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rightArrowButton.Font = New-Object System.Drawing.Font("メイリオ", 12, [System.Drawing.FontStyle]::Bold)
$rightArrowButton.Add_Click({
    if ($script:currentPage -lt ($script:pages.Count - 1)) {
        $script:currentPage++
        Update-ProcessControls
    }
})
$headerPanel.Controls.Add($rightArrowButton)

# ページラベル
$pageLabel = New-Object System.Windows.Forms.Label
$pageLabel.Location = New-Object System.Drawing.Point(420, 10)
$pageLabel.Size = New-Object System.Drawing.Size(200, 30)
$pageLabel.Text = "ページ 1 / $($script:pages.Count)"
$pageLabel.Font = New-Object System.Drawing.Font("メイリオ", 10)
$pageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$headerPanel.Controls.Add($pageLabel)
$script:pageLabel = $pageLabel

# プロセス制御エリア（黄色/ベージュ背景）
$processPanel = New-Object System.Windows.Forms.Panel
$processPanel.Location = New-Object System.Drawing.Point(0, 50)
$processPanel.Size = New-Object System.Drawing.Size(800, 280)
$processPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($processPanel)
$script:processPanel = $processPanel

# ログ表示エリア
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 340)
$logTextBox.Size = New-Object System.Drawing.Size(780, 220)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.ReadOnly = $true
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logTextBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($logTextBox)
$script:logTextBox = $logTextBox

# プロセスコントロールの初期化
Update-ProcessControls

# 初期メッセージ
Write-Log "アプリケーションを起動しました" "INFO"
Write-Log "設定ファイル: $configPath" "INFO"
Write-Log "ページ数: $($script:pages.Count)" "INFO"

# フォームを表示
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
