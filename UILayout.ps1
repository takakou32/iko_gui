# PowerShellスクリプト - UIレイアウト作成
# エンコーディング: UTF-8 BOM付

# GUIフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "プロセス実行GUI"
$form.Size = New-Object System.Drawing.Size(900, 1000)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

# ヘッダー部分（水色背景）
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(900, 50)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
$form.Controls.Add($headerPanel)

# タイトルラベル
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
# 初期タイトルの設定（ページJSONから読み込む）
$initialPageTitle = ""
if ($script:pages.Count -gt 0) {
    $initialPageConfig = $script:pages[0]
    if ($initialPageConfig.JsonPath) {
        $initialPageJsonPath = if ([System.IO.Path]::IsPathRooted($initialPageConfig.JsonPath)) {
            $initialPageConfig.JsonPath
        } else {
            Join-Path $PSScriptRoot $initialPageConfig.JsonPath
        }
        if (Test-Path $initialPageJsonPath) {
            try {
                $initialPageJson = Get-Content $initialPageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
                if ($initialPageJson.Title) {
                    $initialPageTitle = $initialPageJson.Title
                }
            } catch {
                # エラー時は後続のフォールバック処理に任せる
            }
        }
    }
    # フォールバック
    if (-not $initialPageTitle) {
        $initialPageTitle = if ($initialPageConfig.Title) { $initialPageConfig.Title } else { if ($script:config.Title) { $script:config.Title } else { "1.V1 移行ツール適用" } }
    }
} else {
    $initialPageTitle = if ($script:config.Title) { $script:config.Title } else { "1.V1 移行ツール適用" }
}
$titleLabel.Text = $initialPageTitle
$titleLabel.Font = New-Object System.Drawing.Font("メイリオ", 12, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($titleLabel)
$script:titleLabel = $titleLabel

# 左矢印ボタン
$leftArrowButton = New-Object System.Windows.Forms.Button
$leftArrowButton.Location = New-Object System.Drawing.Point(690, 10)
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
$rightArrowButton.Location = New-Object System.Drawing.Point(740, 10)
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
$pageLabel.Size = New-Object System.Drawing.Size(150, 30)
$pageLabel.Text = "ページ 1 / $($script:pages.Count)"
$pageLabel.Font = New-Object System.Drawing.Font("メイリオ", 10)
$pageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$headerPanel.Controls.Add($pageLabel)
$script:pageLabel = $pageLabel

# 編集モード切り替えボタン
$editModeButton = New-Object System.Windows.Forms.Button
$editModeButton.Location = New-Object System.Drawing.Point(580, 10)
$editModeButton.Size = New-Object System.Drawing.Size(100, 30)
$editModeButton.Text = "編集モード OFF"
$editModeButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$editModeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$editModeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$editModeButton.FlatAppearance.BorderSize = 1
$editModeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
$editModeButton.Add_Click({
    $script:editMode = -not $script:editMode
    if ($script:editMode) {
        $editModeButton.Text = "編集モード ON"
        $editModeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
        Write-Log "編集モードを有効にしました" "INFO"
    } else {
        $editModeButton.Text = "編集モード OFF"
        $editModeButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        Write-Log "編集モードを無効にしました" "INFO"
    }
    # ボタンのテキストを更新
    Update-ProcessControls
})
$headerPanel.Controls.Add($editModeButton)
$script:editModeButton = $editModeButton

# プロセス制御エリア（黄色/ベージュ背景）
$processPanel = New-Object System.Windows.Forms.Panel
$processPanel.Location = New-Object System.Drawing.Point(0, 50)
$processPanel.Size = New-Object System.Drawing.Size(900, 320)
$processPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($processPanel)
$script:processPanel = $processPanel

# ファイル移動セクション（黄色/ベージュ背景）
$fileMovePanel = New-Object System.Windows.Forms.Panel
$fileMovePanel.Location = New-Object System.Drawing.Point(0, 370)
$fileMovePanel.Size = New-Object System.Drawing.Size(900, 120)
$fileMovePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($fileMovePanel)

# 移行データファイル移動元ラベル
$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Location = New-Object System.Drawing.Point(10, 10)
$sourceLabel.Size = New-Object System.Drawing.Size(200, 20)
$sourceLabel.Text = "移行データファイル移動元"
$sourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 9)
$fileMovePanel.Controls.Add($sourceLabel)

# 移行データファイル移動元パス入力
$sourcePathTextBox = New-Object System.Windows.Forms.TextBox
$sourcePathTextBox.Location = New-Object System.Drawing.Point(10, 35)
$sourcePathTextBox.Size = New-Object System.Drawing.Size(350, 25)
$sourcePathTextBox.Text = "パス"
$sourcePathTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
$sourcePathTextBox.ReadOnly = $true
$sourcePathTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
$sourcePathTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "移行データファイル移動元フォルダを選択してください"
        $folderDialog.ShowNewFolderButton = $true
        
        # 現在のパスを初期値として設定
        if ($sourcePathTextBox.Text -and $sourcePathTextBox.Text -ne "パス" -and (Test-Path $sourcePathTextBox.Text)) {
            $folderDialog.SelectedPath = $sourcePathTextBox.Text
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            $sourcePathTextBox.Text = $selectedPath
            Save-PagePaths -SourcePath $selectedPath
            Write-Log "移行データファイル移動元を設定しました: $selectedPath" "INFO"
        }
        $folderDialog.Dispose()
    }
})
$fileMovePanel.Controls.Add($sourcePathTextBox)
$script:sourcePathTextBox = $sourcePathTextBox

# 移行データファイル移動先ラベル
$destLabel = New-Object System.Windows.Forms.Label
$destLabel.Location = New-Object System.Drawing.Point(380, 10)
$destLabel.Size = New-Object System.Drawing.Size(200, 20)
$destLabel.Text = "移行データファイル移動先"
$destLabel.Font = New-Object System.Drawing.Font("メイリオ", 9)
$fileMovePanel.Controls.Add($destLabel)

# 移行データファイル移動先パス入力
$destPathTextBox = New-Object System.Windows.Forms.TextBox
$destPathTextBox.Location = New-Object System.Drawing.Point(380, 35)
$destPathTextBox.Size = New-Object System.Drawing.Size(350, 25)
$destPathTextBox.Text = "パス"
$destPathTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
$destPathTextBox.ReadOnly = $true
$destPathTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
$destPathTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "移行データファイル移動先フォルダを選択してください"
        $folderDialog.ShowNewFolderButton = $true
        
        # 現在のパスを初期値として設定
        if ($destPathTextBox.Text -and $destPathTextBox.Text -ne "パス" -and (Test-Path $destPathTextBox.Text)) {
            $folderDialog.SelectedPath = $destPathTextBox.Text
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            $destPathTextBox.Text = $selectedPath
            Save-PagePaths -DestinationPath $selectedPath
            Write-Log "移行データファイル移動先を設定しました: $selectedPath" "INFO"
        }
        $folderDialog.Dispose()
    }
})
$fileMovePanel.Controls.Add($destPathTextBox)
$script:destPathTextBox = $destPathTextBox

# ファイル移動ボタン
$fileMoveButton = New-Object System.Windows.Forms.Button
$fileMoveButton.Location = New-Object System.Drawing.Point(380, 70)
$fileMoveButton.Size = New-Object System.Drawing.Size(120, 35)
$fileMoveButton.Text = "ファイル移動"
$fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 255)
$fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$fileMoveButton.FlatAppearance.BorderSize = 1
$fileMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
$fileMoveButton.Enabled = $false
$fileMovePanel.Controls.Add($fileMoveButton)

# ログ格納セクション（黄色/ベージュ背景）
$logStoragePanel = New-Object System.Windows.Forms.Panel
$logStoragePanel.Location = New-Object System.Drawing.Point(0, 490)
$logStoragePanel.Size = New-Object System.Drawing.Size(900, 80)
$logStoragePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($logStoragePanel)

# ログ格納先ラベル
$logStorageLabel = New-Object System.Windows.Forms.Label
$logStorageLabel.Location = New-Object System.Drawing.Point(10, 10)
$logStorageLabel.Size = New-Object System.Drawing.Size(150, 20)
$logStorageLabel.Text = "ログ格納先"
$logStorageLabel.Font = New-Object System.Drawing.Font("メイリオ", 9)
$logStoragePanel.Controls.Add($logStorageLabel)

# ログ格納先パス入力
$logStoragePathTextBox = New-Object System.Windows.Forms.TextBox
$logStoragePathTextBox.Location = New-Object System.Drawing.Point(10, 35)
$logStoragePathTextBox.Size = New-Object System.Drawing.Size(600, 25)
$logStoragePathTextBox.Text = "パス"
$logStoragePathTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
$logStoragePathTextBox.ReadOnly = $true
$logStoragePathTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
$logStoragePathTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "ログ格納先フォルダを選択してください"
        $folderDialog.ShowNewFolderButton = $true
        
        # 現在のパスを初期値として設定
        if ($logStoragePathTextBox.Text -and $logStoragePathTextBox.Text -ne "パス" -and (Test-Path $logStoragePathTextBox.Text)) {
            $folderDialog.SelectedPath = $logStoragePathTextBox.Text
        } elseif (Test-Path $logDir) {
            $folderDialog.SelectedPath = $logDir
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            $logStoragePathTextBox.Text = $selectedPath
            Save-PagePaths -LogStoragePath $selectedPath
            Write-Log "ログ格納先を設定しました: $selectedPath" "INFO"
        }
        $folderDialog.Dispose()
    }
})
$logStoragePanel.Controls.Add($logStoragePathTextBox)
$script:logStoragePathTextBox = $logStoragePathTextBox

# ログ格納ボタン
$logStorageButton = New-Object System.Windows.Forms.Button
$logStorageButton.Location = New-Object System.Drawing.Point(620, 35)
$logStorageButton.Size = New-Object System.Drawing.Size(100, 25)
$logStorageButton.Text = "ログ格納"
$logStorageButton.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 150)
$logStorageButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$logStorageButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$logStorageButton.FlatAppearance.BorderSize = 1
$logStorageButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
$logStorageButton.Enabled = $false
$logStoragePanel.Controls.Add($logStorageButton)

# ログ出力(全ファイル)ラベル
$logOutputLabel = New-Object System.Windows.Forms.Label
$logOutputLabel.Location = New-Object System.Drawing.Point(10, 580)
$logOutputLabel.Size = New-Object System.Drawing.Size(200, 20)
$logOutputLabel.Text = "ログ出力(全ファイル)"
$logOutputLabel.Font = New-Object System.Drawing.Font("メイリオ", 9)
$form.Controls.Add($logOutputLabel)

# ログ表示エリア
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 605)
$logTextBox.Size = New-Object System.Drawing.Size(880, 220)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.ReadOnly = $true
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logTextBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($logTextBox)
$script:logTextBox = $logTextBox

# プロセスコントロールの初期化
Update-ProcessControls

# 初期ページパスの読み込み（Update-ProcessControls内でLoad-PagePathsが呼ばれるが、念のため）
Load-PagePaths

# 初期メッセージ
Write-Log "アプリケーションを起動しました" "INFO"
Write-Log "設定ファイル: $script:configPath" "INFO"
Write-Log "ページ数: $($script:pages.Count)" "INFO"

# フォームを表示
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
