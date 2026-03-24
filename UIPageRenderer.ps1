# ==============================================================================
# UI描画専用モジュール (UIPageRenderer.ps1)
# 
# 用途: 
#   Functions.ps1からUIの構築処理（Update-ProcessControls等）や
#   コントロール生成処理を分離し、画面のレイアウト・描画のみを担当します。
# 
# ==============================================================================

Write-Host "UIPageRenderer.ps1 loaded."

# ==============================================================================
# コントロール生成補助関数群
# ==============================================================================

# --- 共通ヘルパー関数: メンテボタン作成 ---
function Create-MaintButton {
    param($x, $text, $batchIndex, $title, $componentKey = "MaintButton_Executed")
    $maintButton = New-Object System.Windows.Forms.Button
    $maintButton.Location = New-Object System.Drawing.Point($x, $buttonY)
    $maintButton.Size = New-Object System.Drawing.Size(80, 30)

    if ($script:editMode) {
        $maintButton.Text = "参照"
    }
    else {
        $maintButton.Text = $text
    }
    $maintButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
    $maintButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $maintButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
    $maintButton.FlatAppearance.BorderSize = 1
    $maintButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
    
    $maintButton.Tag = @{
        ProcessIndex = $i
        BatchIndex   = $batchIndex
        Title        = $title
        ComponentKey = $componentKey
    }
    
    $maintButton.Add_Click({
            $ctx = $this.Tag
            $clickedProcessIdx = $ctx.ProcessIndex
            $targetBatchIdx = $ctx.BatchIndex
            $targetTitle = $ctx.Title
            $cKey = $ctx.ComponentKey

            if ($script:editMode) {
                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                $fileDialog.Title = "${targetTitle}用バッチファイルを選択してください"
                
                $pageConfig = $script:pages[$script:currentPage]
                $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                if ($logStoragePath -and (Test-Path $logStoragePath)) {
                    $fileDialog.InitialDirectory = $logStoragePath
                }
                
                # 初期値設定
                $currentProcesses = Get-CurrentPageProcesses
                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                    $processConfig = $currentProcesses[$clickedProcessIdx]
                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt $targetBatchIdx) {
                        $currentBatch = $processConfig.BatchFiles[$targetBatchIdx]
                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                        # パスが有効かチェック（空文字でのTest-Pathエラー防止）
                        if ($initialPath -and (Test-Path $initialPath)) {
                            $fileDialog.InitialDirectory = Split-Path $initialPath
                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                        }
                    }
                    elseif ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                        $currentBatch = $processConfig.BatchFiles[0]
                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                        # パスが有効かチェック
                        if ($initialPath -and (Test-Path $initialPath)) {
                            $fileDialog.InitialDirectory = Split-Path $initialPath
                        }
                    }
                }
                
                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $selectedFile = $fileDialog.FileName
                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex $targetBatchIdx) {
                        Write-Log "${targetTitle}用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                        [System.Windows.Forms.MessageBox]::Show("${targetTitle}用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Update-ProcessControls
                    }
                }
                $fileDialog.Dispose()
            }
            else {
                # 実行
                $currentProcesses = Get-CurrentPageProcesses
                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                    $processConfig = $currentProcesses[$clickedProcessIdx]
                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt $targetBatchIdx) {
                        $batch = $processConfig.BatchFiles[$targetBatchIdx]
                        $batchPath = Resolve-BatchPath -Path $batch.Path
                        
                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey $cKey
                        Update-ProcessControls
                    }
                    else {
                        Write-Log "${targetTitle}用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                        [System.Windows.Forms.MessageBox]::Show("${targetTitle}用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                }
            }
        })
    # 編集モードOFF時：Enabledフラグをmaintボタンに反映
    if (-not $script:editMode -and -not $isEnabled) {
        $maintButton.Enabled = $false
    }
    $script:processPanel.Controls.Add($maintButton)
    return $maintButton
}

# 取込後ボタン作成ヘルパー (BatchIndex可変対応)
function Create-AfterImportButton {
    param($x, $batchIndex, $text, $componentKey = "AfterImportButton_Executed")
    $afterBtn = New-Object System.Windows.Forms.Button
    $afterBtn.Location = New-Object System.Drawing.Point($x, $buttonY)
    $afterBtn.Size = New-Object System.Drawing.Size(80, 30)
    if ($script:editMode) { $afterBtn.Text = "参照" } else { $afterBtn.Text = $text }
    $afterBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)
    $afterBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $afterBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)
    $afterBtn.FlatAppearance.BorderSize = 1
    $afterBtn.Font = New-Object System.Drawing.Font("メイリオ", 9)
    
    $afterBtn.Tag = @{ ProcessIndex = $i; BatchIndex = $batchIndex; Title = $text; ComponentKey = $componentKey }
    $afterBtn.Add_Click({
            $ctx = $this.Tag
            $pIdx = $ctx.ProcessIndex
            $bIdx = $ctx.BatchIndex
            $t = $ctx.Title
            $cKey = $ctx.ComponentKey
            
            if ($script:editMode) {
                # 簡略化のためCreate-MaintButtonと同じロジックを使用
                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                $fileDialog.Title = "${t}用バッチファイルを選択してください"
            
                $pageConfig = $script:pages[$script:currentPage]
                $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                if ($logStoragePath -and (Test-Path $logStoragePath)) { $fileDialog.InitialDirectory = $logStoragePath }
            
                $currentProcesses = Get-CurrentPageProcesses
                if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                    $processConfig = $currentProcesses[$pIdx]
                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt $bIdx) {
                        $currentBatch = $processConfig.BatchFiles[$bIdx]
                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                        if ($initialPath -and (Test-Path $initialPath)) {
                            $fileDialog.InitialDirectory = Split-Path $initialPath
                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                        }
                    }
                    elseif ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                        $currentBatch = $processConfig.BatchFiles[0]
                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                        if ($initialPath -and (Test-Path $initialPath)) { $fileDialog.InitialDirectory = Split-Path $initialPath }
                    }
                }
            
                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $selectedFile = $fileDialog.FileName
                    if (Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $selectedFile -BatchIndex $bIdx) {
                        Write-Log "${t}用バッチファイルを設定しました: $selectedFile" "INFO" $pIdx
                        [System.Windows.Forms.MessageBox]::Show("${t}用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Update-ProcessControls
                    }
                }
                $fileDialog.Dispose()
            }
            else {
                $currentProcesses = Get-CurrentPageProcesses
                if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                    $processConfig = $currentProcesses[$pIdx]
                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt $bIdx) {
                        $batch = $processConfig.BatchFiles[$bIdx]
                        $batchPath = Resolve-BatchPath -Path $batch.Path
                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $pIdx
                        Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey $cKey
                        Update-ProcessControls
                    }
                    else {
                        Write-Log "${t}用バッチファイルが設定されていません" "ERROR" $pIdx
                        [System.Windows.Forms.MessageBox]::Show("${t}用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                }
            }
        })
    # 編集モードOFF時：EnabledフラグをafterBtnに反映
    if (-not $script:editMode -and -not $isEnabled) {
        $afterBtn.Enabled = $false
    }
    $script:processPanel.Controls.Add($afterBtn)
    return $afterBtn
}

# ==============================================================================
# UI描画メイン関数
# ==============================================================================

function Update-ProcessControls {
    # ページ遷移時にJSONファイルを読み込む
    $currentProcesses = @(Get-CurrentPageProcesses)
    $totalPages = $script:pages.Count
    
    Write-Log "ページ $($script:currentPage + 1) のプロセスを読み込みました (プロセス数: $($currentProcesses.Count))" "INFO"
    
    # 既存のコントロールをすべてクリア（processPanel内のすべてのコントロールを削除）
    # まず、processControls配列に保存されているコントロールを削除
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup) {
            if ($ctrlGroup.CheckBox) { $script:processPanel.Controls.Remove($ctrlGroup.CheckBox) }
            if ($ctrlGroup.NameTextBox) { $script:processPanel.Controls.Remove($ctrlGroup.NameTextBox) }
            if ($ctrlGroup.PathTextBox) { $script:processPanel.Controls.Remove($ctrlGroup.PathTextBox) }
            if ($ctrlGroup.KdlSourceTextBox) { $script:processPanel.Controls.Remove($ctrlGroup.KdlSourceTextBox) }
            if ($ctrlGroup.KdlSourceMoveButton) { $script:processPanel.Controls.Remove($ctrlGroup.KdlSourceMoveButton) }
            if ($ctrlGroup.KdlDestTextBox) { $script:processPanel.Controls.Remove($ctrlGroup.KdlDestTextBox) }
            if ($ctrlGroup.KdlDestMoveButton) { $script:processPanel.Controls.Remove($ctrlGroup.KdlDestMoveButton) }
            if ($ctrlGroup.V1CsvDestTextBox) { $script:processPanel.Controls.Remove($ctrlGroup.V1CsvDestTextBox) }
            if ($ctrlGroup.V1CsvDestMoveButton) { $script:processPanel.Controls.Remove($ctrlGroup.V1CsvDestMoveButton) }
            if ($ctrlGroup.KdlImportButton) { $script:processPanel.Controls.Remove($ctrlGroup.KdlImportButton) }
            if ($ctrlGroup.DirectImportButton) { $script:processPanel.Controls.Remove($ctrlGroup.DirectImportButton) }
            if ($ctrlGroup.AfterImportButton) { $script:processPanel.Controls.Remove($ctrlGroup.AfterImportButton) }
            if ($ctrlGroup.FileMoveButton) { $script:processPanel.Controls.Remove($ctrlGroup.FileMoveButton) }
            if ($ctrlGroup.CsvConvertButton) { $script:processPanel.Controls.Remove($ctrlGroup.CsvConvertButton) }
            if ($ctrlGroup.ExecuteButton) { $script:processPanel.Controls.Remove($ctrlGroup.ExecuteButton) }
            if ($ctrlGroup.LogButton) { $script:processPanel.Controls.Remove($ctrlGroup.LogButton) }
        }
    }
    $script:processControls = @()
    
    # 3ページ目・4ページ目の場合、V1抽出CSV格納元・格納先セクションをクリア
    if ($script:currentPage -eq 2 -or $script:currentPage -eq 3) {
        if ($script:v1CsvSourceLabel) { 
            $script:processPanel.Controls.Remove($script:v1CsvSourceLabel)
            $script:v1CsvSourceLabel = $null
        }
        if ($script:v1CsvSourceTextBox) { 
            $script:processPanel.Controls.Remove($script:v1CsvSourceTextBox)
            $script:v1CsvSourceTextBox = $null
        }
        if ($script:v1CsvDestLabel) { 
            $script:processPanel.Controls.Remove($script:v1CsvDestLabel)
            $script:v1CsvDestLabel = $null
        }
    }
    
    # 4ページ目の場合、V1抽出CSV格納元セクションをクリア（既に処理済みの場合はスキップ）
    if ($script:currentPage -eq 3) {
        if ($script:v1CsvSourceLabel) { 
            $script:processPanel.Controls.Remove($script:v1CsvSourceLabel)
            $script:v1CsvSourceLabel = $null
        }
        if ($script:v1CsvSourceTextBox) { 
            $script:processPanel.Controls.Remove($script:v1CsvSourceTextBox)
            $script:v1CsvSourceTextBox = $null
        }
    }
    
    # processPanel内の残っているすべてのコントロールを削除
    # コレクションを反復処理しながら削除すると問題が起きるため、一度配列にコピーしてから削除
    $controlsToRemove = @()
    foreach ($control in $script:processPanel.Controls) {
        $controlsToRemove += $control
    }
    foreach ($control in $controlsToRemove) {
        try {
            $script:processPanel.Controls.Remove($control)
            if ($control -is [System.IDisposable]) {
                $control.Dispose()
            }
        }
        catch {
            # エラーは無視（既に削除されている可能性がある）
        }
    }
    
    # 念のため、processPanel.Controlsをクリア（すべてのコントロールを削除）
    # これにより、前のページのコントロールが確実に削除される
    $script:processPanel.Controls.Clear()
    
    # 4ページ目の背景Paintハンドラをいったん解除（ページが変わる場合も対応）
    if ($script:page4PaintHandler) {
        try { $script:processPanel.remove_Paint($script:page4PaintHandler) } catch {}
        $script:page4PaintHandler = $null
    }
    
    # ページ番号を判定（drawioのレイアウトを適用）
    $isPage1 = ($script:currentPage -eq 0)
    $isPage2 = ($script:currentPage -eq 1)
    $isPage3 = ($script:currentPage -eq 2)
    $isPage4 = ($script:currentPage -eq 3)
    $useDrawioLayout = ($isPage1 -or $isPage2)
    
    # 新しいコントロールを作成
    for ($i = 0; $i -lt $script:processesPerPage; $i++) {
        # 変数の初期化（前回のループの変数が残らないようにする）
        $checkBox = $null
        $nameTextBox = $null
        $v1CsvSourceLabel = $null
        $v1CsvSourceTextBox = $null
        $v1CsvDestLabel = $null
        $v1CsvDestTextBox = $null
        $v1CsvDestMoveButton = $null
        $kdlSourceLabel = $null
        $kdlSourceTextBox = $null
        $kdlSourceMoveButton = $null
        $kdlDestLabel = $null
        $kdlDestTextBox = $null
        $kdlDestMoveButton = $null
        $kdlImportButton = $null
        $directImportButton = $null
        $afterImportButton = $null
        $logButton = $null
        $fileMoveButton = $null
        $csvConvertButton = $null
        $executeButton = $null

        if ($i -lt $currentProcesses.Count) {
            $processConfig = $currentProcesses[$i]
            
            # 有効状態の取得（デフォルトtrue）
            $isEnabled = $true
            if ($processConfig.PSObject.Properties['Enabled']) {
                $isEnabled = $processConfig.Enabled
            }
            
            if ($useDrawioLayout) {
                # ページ1・2：2列レイアウト
                $row = [Math]::Floor($i / 2)
                $col = $i % 2
                # 1ページ目・2ページ目：drawioのレイアウトに合わせる
                # drawioの座標: タスク名(60, 100+), セット/チェック(200, 100+), 実行(280, 100+), ログ確認(350, 100+)
                # プロセスパネルのy座標は50なので、実際のy座標は50から（100-50=50）
                $x = if ($col -eq 0) { 60 } else { 440 }
                $y = 50 + $row * 60
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                # チェックボックス（編集モードON時のみ表示）
                $checkBox = New-Object System.Windows.Forms.CheckBox
                $checkBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 5))
                $checkBox.Size = New-Object System.Drawing.Size(20, 20)
                $checkBox.Visible = $script:editMode
                $script:processPanel.Controls.Add($checkBox)

                # 有効/無効切り替え用チェックボックス（新規）
                $enableCheckBox = New-Object System.Windows.Forms.CheckBox
                $enableCheckBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 30))
                $enableCheckBox.Size = New-Object System.Drawing.Size(20, 20)
                $enableCheckBox.Checked = $isEnabled
                $enableCheckBox.Visible = $script:editMode
                $enableCheckBox.Tag = $i
                $enableCheckBox.Add_Click({
                        $idx = $this.Tag
                        $enabled = $this.Checked
                        Save-ProcessEnabled -ProcessIndex $idx -Enabled $enabled
                        Update-ProcessControls
                    })
                $script:processPanel.Controls.Add($enableCheckBox)
                
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = -not $script:editMode
                $nameTextBox.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 255)
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
                $nameTextBox.Tag = $i
                $nameTextBox.Add_Leave({
                        if ($script:editMode) {
                            $processIdx = $this.Tag
                            $newName = $this.Text
                            Save-ProcessName -ProcessIndex $processIdx -ProcessName $newName
                        }
                    })
                $script:processPanel.Controls.Add($nameTextBox)
                
                # ファイル移動設定ボタン（セット/チェックボタン、赤色）
                # 1ページ目は「実行」ボタンと同じ機能、2ページ目は「セット」（ファイル移動設定）
                $fileMoveButton = New-Object System.Windows.Forms.Button
                $fileMoveX = $x + 140
                $fileMoveButton.Location = New-Object System.Drawing.Point($fileMoveX, $y)
                $fileMoveButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($isPage1) {
                    # 1ページ目：実行ボタンと同じ機能だが、見た目は設計書通りの「チェック」ボタン（fillColor=#ffcccc, strokeColor=#b85450）
                    # 編集モードONの時は「参照」、OFFの時は「チェック」と表示
                    if ($script:editMode) {
                        $fileMoveButton.Text = "参照"
                    }
                    else {
                        $fileMoveButton.Text = "チェック"  # 設計書通り「チェック」と表示
                    }
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc（設計書通りの赤色）
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450（設計書通りの赤色ボーダー）
                    $fileMoveButton.Visible = $true  # 常に表示
                    $fileMoveButton.Tag = $i  # プロセスインデックスをTagに保存
                    $fileMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            
                            if ($script:editMode) {
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "チェック用バッチファイルを選択してください"
                            
                                # 現在のバッチファイルパスを初期値として設定（Index 0）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        $currentBatch = $processConfig.BatchFiles[0]
                                        $initialPath = if ([System.IO.Path]::IsPathRooted($currentBatch.Path)) {
                                            $currentBatch.Path
                                        }
                                        else {
                                            Join-Path $PSScriptRoot $currentBatch.Path
                                        }
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                                        }
                                    }
                                }
                            
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $selectedFile = $fileDialog.FileName
                                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 0) {
                                        Write-Log "チェック用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("チェック用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                        # コントロールを更新して新しい設定を反映
                                        Update-ProcessControls
                                    }
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 編集モードOFF時はチェック用バッチファイルを実行（Index 0）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        $batch = $processConfig.BatchFiles[0]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $batchPath = Resolve-BatchPath -Path $batch.Path
                                        
                                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "FileMoveButton_Executed"
                                        Update-ProcessControls
                                    }
                                    else {
                                        Write-Log "チェック用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("チェック用バッチファイルが設定されていません。`n編集モードで設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                    }
                                }
                            }
                        })
                }
                else {
                    # 2ページ目：セットボタン
                    # 編集モードOFF時は「セット」と表示し、実行ボタンと同じ機能（プロセス実行）
                    # 編集モードON時は「参照」と表示し、実行ボタンの編集モードON時と同じ機能（ファイル選択ウィザードを開き、パスをJSONに保存）
                    if ($script:editMode) {
                        $fileMoveButton.Text = "参照"
                    }
                    else {
                        $fileMoveButton.Text = "セット"
                    }
                    $processIdx = $i
                    $fileMoveButton.Tag = $i  # プロセスインデックスをTagに保存
                    $fileMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag  # Tagからプロセスインデックスを取得
                            # 編集モードON時は実行ボタンと同じ機能（ファイル選択ウィザードを開き、パスをJSONに保存）
                            if ($script:editMode) {
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "チェック用バッチファイルを選択してください"
                            
                                # 現在のバッチファイルパスを初期値として設定（Index 0）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        $currentBatch = $processConfig.BatchFiles[0]
                                        $initialPath = if ([System.IO.Path]::IsPathRooted($currentBatch.Path)) {
                                            $currentBatch.Path
                                        }
                                        else {
                                            Join-Path $PSScriptRoot $currentBatch.Path
                                        }
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                                        }
                                    }
                                }
                            
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $selectedFile = $fileDialog.FileName
                                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 0) {
                                        Write-Log "チェック用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("チェック用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                        # コントロールを更新して新しい設定を反映
                                        Update-ProcessControls
                                    }
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 編集モードOFF時はチェック用バッチファイルを実行（Index 0）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        $batch = $processConfig.BatchFiles[0]
                                        $batch = $processConfig.BatchFiles[0]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $batchPath = Resolve-BatchPath -Path $batch.Path
                                        
                                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "FileMoveButton_Executed"
                                        Update-ProcessControls
                                    }
                                    else {
                                        Write-Log "チェック用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("チェック用バッチファイルが設定されていません。`n編集モードで設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                    }
                                }
                            }
                        })
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450
                    $fileMoveButton.Visible = $true  # 編集モードOFF時も表示
                }
                $fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $fileMoveButton.FlatAppearance.BorderSize = 1
                $fileMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                # 編集モードOFF時：Enabledフラグをボタンに反映
                if (-not $script:editMode -and -not $isEnabled) {
                    $fileMoveButton.Enabled = $false
                }
                $script:processPanel.Controls.Add($fileMoveButton)
                
                # 実行ボタン（オレンジ）
                $executeButton = New-Object System.Windows.Forms.Button
                $executeX = $x + 220
                $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
                $executeButton.Size = New-Object System.Drawing.Size(60, 30)
                if ($script:editMode) {
                    $executeButton.Text = "参照"
                }
                else {
                    $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                }
                # 実行ボタンの見た目（設計書通り：fillColor=#ffcc99, strokeColor=#d6b656）
                $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $executeButton.FlatAppearance.BorderSize = 1
                $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $executeButton.Tag = $i  # プロセスインデックスをTagに保存
                $executeButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        
                        if ($script:currentPage -eq 0 -or $script:currentPage -eq 1) {
                            if ($script:editMode) {
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "実行用バッチファイルを選択してください"
                                
                                # LogStoragePathを初期ディレクトリに設定
                                $pageConfig = $script:pages[$script:currentPage]
                                $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                if ($logStoragePath -and (Test-Path $logStoragePath)) {
                                    $fileDialog.InitialDirectory = $logStoragePath
                                }
                            
                                # 現在のバッチファイルパスを初期値として設定（Index 1）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                        $currentBatch = $processConfig.BatchFiles[1]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                                        }
                                    }
                                }
                            
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $selectedFile = $fileDialog.FileName
                                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 1) {
                                        Write-Log "実行用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("実行用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                        # コントロールを更新して新しい設定を反映
                                        Update-ProcessControls
                                    }
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 編集モードOFF時は実行用バッチファイルを実行（Index 1）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                        $batch = $processConfig.BatchFiles[1]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $batchPath = Resolve-BatchPath -Path $batch.Path
                                        
                                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "ExecuteButton_Executed"
                                        Update-ProcessControls
                                    }
                                    else {
                                        Write-Log "実行用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("実行用バッチファイルが設定されていません。`n編集モードで設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                    }
                                }
                            }
                        }
                        else {
                            Start-ProcessFlow -ProcessIndex $clickedProcessIdx
                        }
                    })
                # 編集モードOFF時：Enabledフラグをexecuteボタンに反映
                if (-not $script:editMode -and -not $isEnabled) {
                    $executeButton.Enabled = $false
                }
                $script:processPanel.Controls.Add($executeButton)
                
                # ログ確認ボタン（緑）
                $logButton = New-Object System.Windows.Forms.Button
                $logX = $x + 290
                $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                $logButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($script:editMode) {
                    $logButton.Text = "参照"
                }
                else {
                    $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                }
                $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                $logButton.FlatAppearance.BorderSize = 1
                $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $logButton.Tag = $i  # プロセスインデックスをTagに保存
                $logButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        Show-ProcessLog -ProcessIndex $clickedProcessIdx
                    })
                $script:processPanel.Controls.Add($logButton)
                
                # 1ページ目・2ページ目用のコントロール情報を保存
                $script:processControls += @{
                    CheckBox       = $checkBox
                    NameTextBox    = $nameTextBox
                    FileMoveButton = $fileMoveButton
                    ExecuteButton  = $executeButton
                    LogButton      = $logButton
                }
            }
            elseif ($isPage3) {
                # 3ページ目：JAVA移行ツール実行のレイアウト（1列レイアウト）
                $row = $i  # 1列レイアウトなので、行番号はインデックスそのまま
                
                # 最初の行の場合のみ、V1抽出CSV格納元・格納先セクションを表示
                if ($i -eq 0) {
                    # V1抽出CSV格納元ラベル
                    $v1CsvSourceLabel = New-Object System.Windows.Forms.Label
                    $v1CsvSourceLabel.Location = New-Object System.Drawing.Point(60, 60)
                    $v1CsvSourceLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvSourceLabel.Text = "V1抽出CSV格納元"
                    $v1CsvSourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvSourceLabel)
                    $script:v1CsvSourceLabel = $v1CsvSourceLabel
                    
                    # V1抽出CSV格納元パス入力
                    $v1CsvSourceTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvSourceTextBox.Location = New-Object System.Drawing.Point(60, 85)
                    $v1CsvSourceTextBox.Size = New-Object System.Drawing.Size(350, 30)
                    $v1CsvSourceTextBox.Text = "パス"
                    $v1CsvSourceTextBox.ReadOnly = $true
                    $v1CsvSourceTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvSourceTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvSourceTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvSourceTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvSourceTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $v1CsvSourceTextBox.Text -Description "V1抽出CSV格納元フォルダを選択してください"
                                if ($selectedPath) {
                                    $v1CsvSourceTextBox.Text = $selectedPath
                                    # page3.jsonに保存
                                    Save-PagePaths -SourcePath $selectedPath
                                    Write-Log "V1抽出CSV格納元を設定しました: $selectedPath" "INFO"
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvSourceTextBox)
                    $script:v1CsvSourceTextBox = $v1CsvSourceTextBox
                    
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(210, 135)
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    $script:v1CsvDestLabel = $v1CsvDestLabel
                }
                
                # drawioの座標: タスク名(60, 210+), パス(210, 210+), 移動設定(440, 210+), CSV名変換(520, 210+), 実行(610, 210+), ログ確認(680, 210+)
                # プロセスパネルのy座標は50なので、実際のy座標は160から（210-50=160）
                $x = 60
                $y = 160 + $row * 40
                
                # チェックボックス（編集モードON時のみ表示: 消去用など）
                $checkBox = New-Object System.Windows.Forms.CheckBox
                $checkBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 5))
                $checkBox.Size = New-Object System.Drawing.Size(20, 20)
                $checkBox.Visible = $script:editMode
                $script:processPanel.Controls.Add($checkBox)
                
                # 有効/無効切り替え用チェックボックス
                $enableCheckBox = New-Object System.Windows.Forms.CheckBox
                $enableCheckBox.Location = New-Object System.Drawing.Point([int]($x - 45), [int]($y + 5))
                $enableCheckBox.Size = New-Object System.Drawing.Size(20, 20)
                $enableCheckBox.Checked = $isEnabled
                $enableCheckBox.Visible = $script:editMode
                $enableCheckBox.Tag = $i
                $enableCheckBox.Add_Click({
                        $idx = $this.Tag
                        $enabled = $this.Checked
                        Save-ProcessEnabled -ProcessIndex $idx -Enabled $enabled
                        Update-ProcessControls
                    })
                $script:processPanel.Controls.Add($enableCheckBox)
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = -not $script:editMode
                $nameTextBox.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 255)
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
                $nameTextBox.Tag = $i
                $nameTextBox.Add_Leave({
                        if ($script:editMode) {
                            $processIdx = $this.Tag
                            $newName = $this.Text
                            Save-ProcessName -ProcessIndex $processIdx -ProcessName $newName
                        }
                    })
                $script:processPanel.Controls.Add($nameTextBox)
                
                # パス入力テキストボックス（V1抽出CSV格納先）
                $pathTextBox = New-Object System.Windows.Forms.TextBox
                $pathX = 210
                $pathTextBox.Location = New-Object System.Drawing.Point($pathX, $y)
                $pathTextBox.Size = New-Object System.Drawing.Size(220, 30)
                # 各プロセスのDestinationPathを読み込んで設定
                $destPathValue = "パス"
                if ($processConfig.DestinationPath -and $processConfig.DestinationPath -ne "" -and $processConfig.DestinationPath -ne "パス") {
                    try {
                        $destPathValue = $processConfig.DestinationPath
                        # 相対パスの場合は絶対パスに変換
                        if (-not [System.IO.Path]::IsPathRooted($destPathValue)) {
                            # 共通基準パスを使用
                            $basePath = Get-CommonBasePath
                            $destPathValue = Join-Path $basePath $destPathValue
                        }
                        $destPathValue = [System.IO.Path]::GetFullPath($destPathValue)
                    }
                    catch {
                        # エラー時はデフォルト値を使用
                    }
                }
                $pathTextBox.Text = $destPathValue
                $pathTextBox.ReadOnly = $true
                $pathTextBox.BackColor = [System.Drawing.Color]::White
                $pathTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $pathTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $pathTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                $pathTextBox.Tag = $i  # プロセスインデックスをTagに保存
                $pathTextBox.Add_Click({
                        if ($script:editMode) {
                            $initialDir = "パス"
                            if ($this.Text -ne "パス") {
                                $initialDir = $this.Text
                            }
                            $selectedPath = Show-FolderBrowser -InitialDirectory $initialDir -Description "V1抽出CSV格納先フォルダを選択してください"
                            if ($selectedPath) {
                                $this.Text = $selectedPath
                                # 各プロセスのDestinationPathをpage3.jsonに保存
                                $clickedProcessIdx = $this.Tag
                                Save-ProcessDestinationPath -ProcessIndex $clickedProcessIdx -DestinationPath $selectedPath
                                Write-Log "V1抽出CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                            }
                        }
                        else {
                            $path = $this.Text
                            if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                if (-not [System.IO.Path]::IsPathRooted($path)) {
                                    $path = Join-Path $PSScriptRoot $path
                                }
                                if (Test-Path $path) {
                                    Open-PathInExplorer -Path $path
                                }
                                else {
                                    Write-Log "パスが存在しません: $path" "WARN"
                                }
                            }
                        }
                    })
                $script:processPanel.Controls.Add($pathTextBox)
                
                # 移動設定ボタン（編集モードON時は水色、OFF時は紺色）
                $fileMoveButton = New-Object System.Windows.Forms.Button
                $fileMoveX = 440
                $fileMoveButton.Location = New-Object System.Drawing.Point($fileMoveX, $y)
                $fileMoveButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($script:editMode) {
                    $fileMoveButton.Text = "移動設定"
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                }
                else {
                    $fileMoveButton.Text = "移動"
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                }
                $fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $fileMoveButton.FlatAppearance.BorderSize = 1
                $fileMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $fileMoveButton.Visible = $true  # 常に表示
                $fileMoveButton.Tag = $i
                $fileMoveButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        $currentProcessName = ""
                        $v1CsvSourcePath = ""
                        $v1CsvDestPath = ""
                    
                        # プロセス名とパスを取得
                        if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                            $ctrlGroup = $script:processControls[$clickedProcessIdx]
                            if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
                                $currentProcessName = $ctrlGroup.NameTextBox.Text
                            }
                            if ($ctrlGroup -and $ctrlGroup.PathTextBox) {
                                $v1CsvDestPath = $ctrlGroup.PathTextBox.Text
                            }
                        }
                    
                        # V1抽出CSV格納元を取得
                        if ($script:v1CsvSourceTextBox) {
                            $v1CsvSourcePath = $script:v1CsvSourceTextBox.Text
                        }
                    
                        # 編集モードと非編集モードで動作を分岐
                        if ($script:editMode) {
                            # 編集モード：移動設定ダイアログを表示
                            Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName
                        }
                        else {
                            # 非編集モード：ファイル移動を実行
                            Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $v1CsvSourcePath -V1CsvDestinationPath $v1CsvDestPath
                        }
                    })
                $script:processPanel.Controls.Add($fileMoveButton)
                
                # CSV名変換ボタン（赤色）- 実行ボタンと同じ機能
                $csvConvertButton = New-Object System.Windows.Forms.Button
                $csvConvertX = 520
                $csvConvertButton.Location = New-Object System.Drawing.Point($csvConvertX, $y)
                $csvConvertButton.Size = New-Object System.Drawing.Size(80, 30)
                if ($script:editMode) {
                    $csvConvertButton.Text = "参照"
                }
                else {
                    $csvConvertButton.Text = "CSV名変換"
                }
                $csvConvertButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc
                $csvConvertButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $csvConvertButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450
                $csvConvertButton.FlatAppearance.BorderSize = 1
                $csvConvertButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $csvConvertButton.Tag = $i  # プロセスインデックスをTagに保存
                $csvConvertButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        if ($script:editMode) {
                            # 編集モードON：ファイル選択ダイアログでバッチファイルのパスをJSONに保存（CSV名変換用：Index 1）
                            $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                            $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                            $fileDialog.Title = "CSV名変換用バッチファイルを選択してください"
                            
                            # LogStoragePathを初期ディレクトリに設定
                            $pageConfig = $script:pages[$script:currentPage]
                            $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                            if ($logStoragePath -and (Test-Path $logStoragePath)) {
                                $fileDialog.InitialDirectory = $logStoragePath
                            }
                            
                            # 現在のバッチファイルパスを初期値として設定（BatchIndex = 1）
                            $currentProcesses = Get-CurrentPageProcesses
                            if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                $processConfig = $currentProcesses[$clickedProcessIdx]
                                if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                    $currentBatch = $processConfig.BatchFiles[1]
                                    # Resolve-BatchPathを使用してパスを解決
                                    $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                    if (Test-Path $initialPath) {
                                        $fileDialog.InitialDirectory = Split-Path $initialPath
                                        $fileDialog.FileName = Split-Path $initialPath -Leaf
                                    }
                                }
                            }
                            
                            if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $selectedFile = $fileDialog.FileName
                                # Save-BatchFilePath内でパス制限チェックが行われる
                                if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 1) {
                                    Write-Log "CSV名変換用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                    [System.Windows.Forms.MessageBox]::Show("CSV名変換用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                    # コントロールを更新して新しい設定を反映
                                    Update-ProcessControls
                                }
                            }
                            $fileDialog.Dispose()
                        }
                        else {
                            # 編集モードOFF：JSONに設定されたバッチファイルを実行（BatchIndex = 1）
                            $currentProcesses = Get-CurrentPageProcesses
                            if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                $processConfig = $currentProcesses[$clickedProcessIdx]
                                if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                    $batch = $processConfig.BatchFiles[1]
                                    # Resolve-BatchPathを使用してパスを解決
                                    $batchPath = Resolve-BatchPath -Path $batch.Path
                                    
                                    $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                    Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "CsvConvertButton_Executed"
                                    Update-ProcessControls
                                }
                                else {
                                    Write-Log "CSV名変換用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                    [System.Windows.Forms.MessageBox]::Show("CSV名変換用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                }
                            }
                        }
                    })
                # 編集モードOFF時：Enabledフラグをcsvconvertボタンに反映
                if (-not $script:editMode -and -not $isEnabled) {
                    $csvConvertButton.Enabled = $false
                }
                $script:processPanel.Controls.Add($csvConvertButton)
                
                # 実行ボタン（オレンジ）
                $executeButton = New-Object System.Windows.Forms.Button
                $executeX = 610
                $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
                $executeButton.Size = New-Object System.Drawing.Size(60, 30)
                if ($script:editMode) {
                    $executeButton.Text = "参照"
                }
                else {
                    $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                }
                $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                $executeButton.FlatAppearance.BorderSize = 1
                $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $executeButton.Tag = $i  # プロセスインデックスをTagに保存
                $executeButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        if ($script:editMode) {
                            # 編集モードON：ファイル選択ダイアログでバッチファイルのパスをJSONに保存（実行用：Index 0）
                            $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                            $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                            $fileDialog.Title = "実行用バッチファイルを選択してください"
                            
                            # LogStoragePathを初期ディレクトリに設定
                            $pageConfig = $script:pages[$script:currentPage]
                            $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                            if ($logStoragePath -and (Test-Path $logStoragePath)) {
                                $fileDialog.InitialDirectory = $logStoragePath
                            }
                            
                            # 現在のバッチファイルパスを初期値として設定（BatchIndex = 0）
                            $currentProcesses = Get-CurrentPageProcesses
                            if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                $processConfig = $currentProcesses[$clickedProcessIdx]
                                if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                    $currentBatch = $processConfig.BatchFiles[0]
                                    # Resolve-BatchPathを使用してパスを解決
                                    $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                    if (Test-Path $initialPath) {
                                        $fileDialog.InitialDirectory = Split-Path $initialPath
                                        $fileDialog.FileName = Split-Path $initialPath -Leaf
                                    }
                                }
                            }
                            
                            if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $selectedFile = $fileDialog.FileName
                                if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 0) {
                                    Write-Log "実行用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                    [System.Windows.Forms.MessageBox]::Show("実行用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                    # コントロールを更新して新しい設定を反映
                                    Update-ProcessControls
                                }
                            }
                            $fileDialog.Dispose()
                        }
                        else {
                            # 編集モードOFF：JSONに設定されたバッチファイルを実行（BatchIndex = 0）
                            $currentProcesses = Get-CurrentPageProcesses
                            if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                $processConfig = $currentProcesses[$clickedProcessIdx]
                                if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                    $batch = $processConfig.BatchFiles[0]
                                    # Resolve-BatchPathを使用してパスを解決
                                    $batchPath = Resolve-BatchPath -Path $batch.Path
                                    $this.Enabled = $false
                                    $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                    $this.Enabled = $true
                                    Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "ExecuteButton_Executed"
                                    Update-ProcessControls
                                }
                                else {
                                    Write-Log "実行用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                    [System.Windows.Forms.MessageBox]::Show("実行用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                }
                            }
                        }
                    })
                $script:processPanel.Controls.Add($executeButton)
                
                # ログ確認ボタン（緑）
                $logButton = New-Object System.Windows.Forms.Button
                $logX = 680
                $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                $logButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($script:editMode) {
                    $logButton.Text = "参照"
                }
                else {
                    $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                }
                $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                $logButton.FlatAppearance.BorderSize = 1
                $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $logButton.Tag = $i  # プロセスインデックスをTagに保存
                $logButton.Add_Click({
                        $clickedProcessIdx = $this.Tag
                        Show-ProcessLog -ProcessIndex $clickedProcessIdx
                    })
                $script:processPanel.Controls.Add($logButton)
                
                # 3ページ目用のコントロール情報を保存
                $script:processControls += @{
                    CheckBox         = $checkBox
                    EnableCheckBox   = $enableCheckBox
                    NameTextBox      = $nameTextBox
                    PathTextBox      = $pathTextBox
                    FileMoveButton   = $fileMoveButton
                    CsvConvertButton = $csvConvertButton
                    ExecuteButton    = $executeButton
                    LogButton        = $logButton
                }
            }
            elseif ($isPage4) {
                # 4ページ目：SQLLOADER実行のレイアウト（1列レイアウト）
                $row = $i  # 1列レイアウトなので、行番号はインデックスそのまま
                
                # プロセス行の座標計算
                $x = 35
                if ($row -lt 2) {
                    $y = [int](140 + $row * 220)
                }
                else {
                    # 3行目以降は行間を詰める
                    $y = [int](580 + ($row - 2) * 80)
                    
                }

                # 4ページ目背景色: 最初の行勦だけ processPanel の Paint イベントを登録する
                if ($i -eq 0) {
                    $script:page4RowColors = @(
                        [System.Drawing.Color]::FromArgb(210, 230, 245),  # 0: スカイブルー
                        [System.Drawing.Color]::FromArgb(215, 238, 215),  # 1: 薄緑
                        [System.Drawing.Color]::FromArgb(100, 100, 240),  # 2: ブルー
                        [System.Drawing.Color]::FromArgb(220, 225, 190),  # 3: 油色
                        [System.Drawing.Color]::FromArgb(100, 180, 100),  # 4: 緑
                        [System.Drawing.Color]::FromArgb(255, 180, 255),  # 5: 赤紫
                        [System.Drawing.Color]::FromArgb(240, 215, 225)   # 6: ピンク
                    )
                    # 既存ハンドラを解除して再登録
                    if ($script:page4PaintHandler) {
                        try { $script:processPanel.remove_Paint($script:page4PaintHandler) } catch {}
                    }
                    $script:page4PaintHandler = [System.Windows.Forms.PaintEventHandler] {
                        param($paintSender, $paintArgs)
                        $g = $paintArgs.Graphics
                        $offsetY = $paintSender.AutoScrollPosition.Y
                        $w = [Math]::Max($paintSender.ClientSize.Width, 850)

                        # 行ごとの矩形 (TopY, Height)
                        # y[0]=140→top=10、y[1]=360→top=335、y[2]=580→top=560
                        # y[3]=660→top=640、y[4]=740→top=720、y[5]=820→top=800、y[6]=900→top=880
                        $rowRects = @(
                            @{Top = 10; H = 325 },  # 0: スカイブルー (y=10~335)
                            @{Top = 335; H = 225 },  # 1: 薄緑       (y=335~560)
                            @{Top = 560; H = 80 },   # 2: ブルー     (y=560~640)
                            @{Top = 640; H = 80 },   # 3: 油色       (y=640~720)
                            @{Top = 720; H = 80 },   # 4: 緑         (y=720~800)
                            @{Top = 800; H = 80 },   # 5: 赤紫       (y=800~880)
                            @{Top = 880; H = 80 }    # 6: ピンク     (y=880~960)
                        )
                        $cols = $script:page4RowColors
                        for ($ri = 0; $ri -lt [Math]::Min($rowRects.Count, $cols.Count); $ri++) {
                            $rect = $rowRects[$ri]
                            $brush = New-Object System.Drawing.SolidBrush($cols[$ri])
                            $g.FillRectangle($brush, 0, ($rect.Top + $offsetY), $w, $rect.H)
                            $brush.Dispose()
                        }
                    }
                    $script:processPanel.add_Paint($script:page4PaintHandler)
                    # 再描画を強制
                    $script:processPanel.Invalidate()
                }

                if ($i -eq 0) {
                    # 最初の行の場合のみ、V1抽出CSV格納元セクションを表示
                    # V1抽出CSV格納元ラベル
                    $v1CsvSourceLabel = New-Object System.Windows.Forms.Label
                    $v1CsvSourceLabel.Location = New-Object System.Drawing.Point(35, 30)
                    $v1CsvSourceLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvSourceLabel.Text = "V1抽出CSV格納元"
                    $v1CsvSourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvSourceLabel)
                    $script:v1CsvSourceLabel = $v1CsvSourceLabel
                    
                    # V1抽出CSV格納元パス入力
                    $v1CsvSourceTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvSourceTextBox.Location = New-Object System.Drawing.Point(35, 50)
                    $v1CsvSourceTextBox.Size = New-Object System.Drawing.Size(360, 30)
                    $v1CsvSourceTextBox.Text = "パス"
                    $v1CsvSourceTextBox.ReadOnly = $true
                    $v1CsvSourceTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvSourceTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvSourceTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvSourceTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvSourceTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $v1CsvSourceTextBox.Text -Description "V1抽出CSV格納元フォルダを選択してください"
                                if ($selectedPath) {
                                    $v1CsvSourceTextBox.Text = $selectedPath
                                    Save-PagePaths -SourcePath $selectedPath
                                    Write-Log "V1抽出CSV格納元を設定しました: $selectedPath" "INFO"
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvSourceTextBox)
                    $script:v1CsvSourceTextBox = $v1CsvSourceTextBox
                }
                
                # タスク名
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = -not $script:editMode
                $nameTextBox.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 255)
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
                $nameTextBox.Tag = $i
                $nameTextBox.Add_Leave({
                        if ($script:editMode) {
                            $processIdx = $this.Tag
                            $newName = $this.Text
                            Save-ProcessName -ProcessIndex $processIdx -ProcessName $newName
                        }
                    })
                $script:processPanel.Controls.Add($nameTextBox)

                # 有効/無効切り替え用チェックボックス（Page 4 Index 0-1 用）
                $enableCheckBox = New-Object System.Windows.Forms.CheckBox
                $enableCheckBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 5))
                $enableCheckBox.Size = New-Object System.Drawing.Size(20, 20)
                $enableCheckBox.Checked = $isEnabled
                $enableCheckBox.Visible = $script:editMode
                $enableCheckBox.Tag = $i
                if ($i -lt 3) {
                    # Add_Click event only for index 0, 1, 2
                    $enableCheckBox.Add_Click({
                            $idx = $this.Tag
                            $enabled = $this.Checked
                            Save-ProcessEnabled -ProcessIndex $idx -Enabled $enabled
                            Update-ProcessControls
                        })
                }
                $script:processPanel.Controls.Add($enableCheckBox)
                
                if ($i -lt 2) {

                    # KDL変換CSV格納元ラベル
                    $kdlSourceLabel = New-Object System.Windows.Forms.Label
                    $kdlSourceLabel.Location = New-Object System.Drawing.Point(175, [int]($y - 20))
                    $kdlSourceLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $kdlSourceLabel.Text = "KDL変換CSV格納元"
                    $kdlSourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($kdlSourceLabel)
                    
                    # KDL変換CSV格納元パス入力
                    $kdlSourceTextBox = New-Object System.Windows.Forms.TextBox
                    $kdlSourceTextBox.Location = New-Object System.Drawing.Point(175, $y)
                    $kdlSourceTextBox.Size = New-Object System.Drawing.Size(260, 30)
                    $kdlSourceTextBox.Text = "パス"
                    $kdlSourceTextBox.ReadOnly = $true
                    $kdlSourceTextBox.BackColor = [System.Drawing.Color]::White
                    $kdlSourceTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $kdlSourceTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $kdlSourceTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $kdlSourceTextBox.Tag = $i  # プロセスインデックスをTagに保存
                    $kdlSourceTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $this.Text -Description "KDL変換CSV格納元フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    # 各プロセスのKdlSourcePathをpage4.jsonに保存
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessKdlSourcePath -ProcessIndex $clickedProcessIdx -KdlSourcePath $selectedPath
                                    Write-Log "KDL変換CSV格納元を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }

                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    # KDL変換CSV格納元の初期値を設定
                    $kdlSourcePathValue = "パス"
                    if ($processConfig.KdlSourcePath -and $processConfig.KdlSourcePath -ne "" -and $processConfig.KdlSourcePath -ne "パス") {
                        try {
                            $kdlSourcePathValue = $processConfig.KdlSourcePath
                            # 相対パスの場合は絶対パスに変換
                            if (-not [System.IO.Path]::IsPathRooted($kdlSourcePathValue)) {
                                $basePath = Get-CommonBasePath
                                $kdlSourcePathValue = Join-Path $basePath $kdlSourcePathValue
                            }
                            $kdlSourcePathValue = [System.IO.Path]::GetFullPath($kdlSourcePathValue)
                        }
                        catch {
                            # エラー時はデフォルト値を使用
                            Write-Log "パスの解決に失敗しました (Process: $i, Type: V1CsvDestPath): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $kdlSourceTextBox.Text = $kdlSourcePathValue
                    $script:processPanel.Controls.Add($kdlSourceTextBox)
                    

                    
                    # KDL変換CSV格納先ラベル
                    $kdlDestLabel = New-Object System.Windows.Forms.Label
                    $kdlDestLabel.Location = New-Object System.Drawing.Point(515, [int]($y - 20))
                    $kdlDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $kdlDestLabel.Text = "KDL変換CSV格納先"
                    $kdlDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($kdlDestLabel)
                    
                    # KDL変換CSV格納先パス入力
                    $kdlDestTextBox = New-Object System.Windows.Forms.TextBox
                    $kdlDestTextBox.Location = New-Object System.Drawing.Point(515, $y)
                    $kdlDestTextBox.Size = New-Object System.Drawing.Size(230, 30)
                    $kdlDestTextBox.Text = "パス"
                    $kdlDestTextBox.ReadOnly = $true
                    $kdlDestTextBox.BackColor = [System.Drawing.Color]::White
                    $kdlDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $kdlDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $kdlDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $kdlDestTextBox.Tag = $i  # プロセスインデックスをTagに保存
                    $kdlDestTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $this.Text -Description "KDL変換CSV格納先フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    # 各プロセスのKdlDestPathをpage4.jsonに保存
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessKdlDestPath -ProcessIndex $clickedProcessIdx -KdlDestPath $selectedPath
                                    Write-Log "KDL変換CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }

                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    # KDL変換CSV格納先の初期値を設定
                    $kdlDestPathValue = "パス"
                    if ($processConfig.KdlDestPath -and $processConfig.KdlDestPath -ne "" -and $processConfig.KdlDestPath -ne "パス") {
                        try {
                            $kdlDestPathValue = $processConfig.KdlDestPath
                            # 相対パスの場合は絶対パスに変換
                            if (-not [System.IO.Path]::IsPathRooted($kdlDestPathValue)) {
                                $basePath = Get-CommonBasePath
                                $kdlDestPathValue = Join-Path $basePath $kdlDestPathValue
                            }
                            $kdlDestPathValue = [System.IO.Path]::GetFullPath($kdlDestPathValue)
                        }
                        catch {
                            # エラー時はデフォルト値を使用
                            Write-Log "パスの解決に失敗しました (Process: $i, Type: V1CsvDestPath): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $kdlDestTextBox.Text = $kdlDestPathValue
                    $script:processPanel.Controls.Add($kdlDestTextBox)
                    
                    # KDL変換CSV格納先の移動設定ボタン（編集モードON時は水色、OFF時は紺色）
                    $kdlDestMoveButton = New-Object System.Windows.Forms.Button
                    $kdlDestMoveButton.Location = New-Object System.Drawing.Point(750, $y)
                    $kdlDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) {
                        $kdlDestMoveButton.Text = "移動設定"
                        $kdlDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                        $kdlDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    }
                    else {
                        $kdlDestMoveButton.Text = "移動"
                        $kdlDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                        $kdlDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                    }
                    $kdlDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $kdlDestMoveButton.FlatAppearance.BorderSize = 1
                    $kdlDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $kdlDestMoveButton.Visible = $true  # 常に表示
                    $kdlDestMoveButton.Tag = $i
                    $kdlDestMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            $kdlSourcePath = ""
                            $kdlDestPath = ""
                        
                            # プロセス名とパスを取得
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
                                    $currentProcessName = $ctrlGroup.NameTextBox.Text
                                }
                                if ($ctrlGroup -and $ctrlGroup.KdlSourceTextBox) {
                                    $kdlSourcePath = $ctrlGroup.KdlSourceTextBox.Text
                                }
                                if ($ctrlGroup -and $ctrlGroup.KdlDestTextBox) {
                                    $kdlDestPath = $ctrlGroup.KdlDestTextBox.Text
                                }
                            }
                        
                            # 編集モードと非編集モードで動作を分岐
                            if ($script:editMode) {
                                # 編集モード：移動設定ダイアログを表示
                                Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -FileSuffix "_1"
                            }
                            else {
                                # 非編集モード：ファイルコピーを実行（KDL格納元 -> KDL格納先）
                                Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $kdlSourcePath -V1CsvDestinationPath $kdlDestPath -FileSuffix "_1" -IsCopy $true
                            }
                        })
                    $script:processPanel.Controls.Add($kdlDestMoveButton)
                    
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(490, [int]($y + 55))
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    
                    # V1抽出CSV格納先パス入力
                    $v1CsvDestTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvDestTextBox.Location = New-Object System.Drawing.Point(490, [int]($y + 75))
                    $v1CsvDestTextBox.Size = New-Object System.Drawing.Size(230, 30)
                    $v1CsvDestTextBox.Text = "パス"
                    $v1CsvDestTextBox.ReadOnly = $true
                    $v1CsvDestTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvDestTextBox.Tag = $i  # プロセスインデックスをTagに保存（1行目・2行目）
                    $v1CsvDestTextBox.Add_Click({
                            if ($script:editMode) {
                                $initialDir = "パス"
                                if ($this.Text -ne "パス") {
                                    $initialDir = $this.Text
                                }
                                $selectedPath = Show-FolderBrowser -InitialDirectory $initialDir -Description "V1抽出CSV格納先フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    # 各プロセスのV1CsvDestPathをpage4.jsonに保存
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessV1CsvDestPath -ProcessIndex $clickedProcessIdx -V1CsvDestPath $selectedPath
                                    Write-Log "V1抽出CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    # V1抽出CSV格納先の初期値を設定（1行目・2行目）
                    $v1CsvDestPathValue = "パス"
                    if ($processConfig.V1CsvDestPath -and $processConfig.V1CsvDestPath -ne "" -and $processConfig.V1CsvDestPath -ne "パス") {
                        try {
                            $v1CsvDestPathValue = $processConfig.V1CsvDestPath
                            # 相対パスの場合は絶対パスに変換
                            if (-not [System.IO.Path]::IsPathRooted($v1CsvDestPathValue)) {
                                # 共通基準パスを使用
                                $basePath = Get-CommonBasePath
                                $v1CsvDestPathValue = Join-Path $basePath $v1CsvDestPathValue
                            }
                            $v1CsvDestPathValue = [System.IO.Path]::GetFullPath($v1CsvDestPathValue)
                        }
                        catch {
                            # エラー時はデフォルト値を使用
                            Write-Log "パスの解決に失敗しました (Process: $i, Type: V1CsvDestPath): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $v1CsvDestTextBox.Text = $v1CsvDestPathValue
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # V1抽出CSV格納先の移動設定ボタン（編集モードON時は水色、OFF時は紺色）
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(725, [int]($y + 75))
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) {
                        $v1CsvDestMoveButton.Text = "移動設定"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    }
                    else {
                        $v1CsvDestMoveButton.Text = "移動"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                    }
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $true  # 常に表示
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            $v1CsvSourcePath = ""
                            $v1CsvDestPath = ""
                        
                            # プロセス名とパスを取得
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
                                    $currentProcessName = $ctrlGroup.NameTextBox.Text
                                }
                                if ($ctrlGroup -and $ctrlGroup.V1CsvDestTextBox) {
                                    $v1CsvDestPath = $ctrlGroup.V1CsvDestTextBox.Text
                                }
                            }
                        
                            # V1抽出CSV格納元を取得
                            if ($script:v1CsvSourceTextBox) {
                                $v1CsvSourcePath = $script:v1CsvSourceTextBox.Text
                            }
                        
                            # 編集モードと非編集モードで動作を分岐
                            if ($script:editMode) {
                                # 編集モード：移動設定ダイアログを表示
                                Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -FileSuffix "_2"
                            }
                            else {
                                # 非編集モード：ファイルコピーを実行（V1格納元 -> V1格納先）
                                Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $v1CsvSourcePath -V1CsvDestinationPath $v1CsvDestPath -FileSuffix "_2" -IsCopy $true
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    # ボタン行
                    $buttonY = [int]($y + 115)
                    
                    # Row 2 (Index 1) 用の標準座標
                    $kdlImportX = 240
                    $directImportX = 340
                    $afterImportX = 440
                    $maint1X = 530
                    $logX = 710
                    
                    if ($i -eq 0) {
                        # --- 1行目 (Index 0) 特殊レイアウト ---
                        # KDL取込(KDB), KDL取込(EB), 直接取込 など
                        $kdlKdbX = 240
                        $kdlEbX = 340
                        $directImportX_Row1 = 440
                        
                        # KDL取込(KDB)ボタン (Batch Index 0)
                        $kdlKdbButton = New-Object System.Windows.Forms.Button
                        $kdlKdbButton.Location = New-Object System.Drawing.Point($kdlKdbX, $buttonY)
                        $kdlKdbButton.Size = New-Object System.Drawing.Size(90, 30)
                        if ($script:editMode) { $kdlKdbButton.Text = "参照" } else { $kdlKdbButton.Text = "KDL取込(KDB)" }
                        $kdlKdbButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)
                        $kdlKdbButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $kdlKdbButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)
                        $kdlKdbButton.FlatAppearance.BorderSize = 1
                        $kdlKdbButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                        $kdlKdbButton.Tag = @{ ProcessIndex = $i; BatchIndex = 0; Title = "KDL取込(KDB)" }
                        $kdlKdbButton.Add_Click({
                                $ctx = $this.Tag
                                $pIdx = $ctx.ProcessIndex
                                $bIdx = $ctx.BatchIndex
                                $title = $ctx.Title
                                if ($script:editMode) {
                                    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                    $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                    $fileDialog.Title = "$title 用バッチファイルを選択してください"
                                    
                                    # 初期ディレクトリの設定
                                    $initDir = ""
                                    
                                    # 既存の設定を確認
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $currentBatch = $procConf.BatchFiles[$bIdx]
                                            if ($currentBatch.Path) {
                                                $resolvedPath = Resolve-BatchPath -Path $currentBatch.Path
                                                if (Test-Path $resolvedPath) {
                                                    $initDir = Split-Path $resolvedPath -Parent
                                                }
                                            }
                                        }
                                    }
                                    
                                    # 既存がない、または無効な場合はGlobalLogPath/LogStoragePath
                                    if (-not $initDir -or -not (Test-Path $initDir)) {
                                        $logStoragePath = if ($script:globalLogPath) { $script:globalLogPath } else { "" }
                                        
                                        if (-not $logStoragePath) {
                                            $pageConfig = $script:pages[$script:currentPage]
                                            $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                            
                                            # メモリ上にない場合、JSONファイルから直接読み込む試み
                                            if (-not $logStoragePath -and $pageConfig.JsonPath) {
                                                $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
                                                    $pageConfig.JsonPath
                                                }
                                                else {
                                                    Join-Path $PSScriptRoot $pageConfig.JsonPath
                                                }
                                                
                                                if (Test-Path $jsonPath) {
                                                    try {
                                                        $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                                                        if ($pageJson.LogStoragePath) {
                                                            $logStoragePath = $pageJson.LogStoragePath
                                                            # メモリ上の設定も更新
                                                            $pageConfig.LogStoragePath = $logStoragePath
                                                        }
                                                    }
                                                    catch {}
                                                }
                                            }
                                        }

                                        if ($logStoragePath -and (Test-Path $initDir)) {
                                            $fileDialog.InitialDirectory = $initDir
                                        }
                                    }
                                    
                                    if ($initDir -and (Test-Path $initDir)) {
                                        $fileDialog.InitialDirectory = $initDir
                                    }

                                    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                        Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $fileDialog.FileName -BatchIndex $bIdx
                                        Update-ProcessControls
                                    }
                                    $fileDialog.Dispose()
                                }
                                else {
                                    # 引数の準備 (KdlSourcePath, KdlDestPath)
                                    $batchArgs = @()
                                    if ($script:processControls -and $pIdx -lt $script:processControls.Count) {
                                        $ctrls = $script:processControls[$pIdx]
                                        $src = if ($ctrls.KdlSourceTextBox) { $ctrls.KdlSourceTextBox.Text } else { "" }
                                        $dst = if ($ctrls.KdlDestTextBox) { $ctrls.KdlDestTextBox.Text } else { "" }
                                        if ($src -eq "パス") { $src = "" }
                                        if ($dst -eq "パス") { $dst = "" }
                                        $batchArgs = @($src, $dst)
                                    }

                                    # Get-BatchFilePathの代わり: インラインでパスを取得して解決
                                    $batchPath = $null
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $batch = $procConf.BatchFiles[$bIdx]
                                            $batchPath = Resolve-BatchPath -Path $batch.Path
                                        }
                                    }

                                    if ($batchPath) {
                                        Invoke-BatchFile -BatchPath $batchPath -DisplayName $title -ProcessIndex $pIdx -Arguments $batchArgs
                                        Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey "KdlKdbButton_Executed"
                                        Update-ProcessControls
                                    }
                                }
                            })
                        # 編集モードOFF時：EnabledフラグをkdlKdbボタンに反映
                        if (-not $script:editMode -and -not $isEnabled) {
                            $kdlKdbButton.Enabled = $false
                        }
                        $script:processPanel.Controls.Add($kdlKdbButton)
                        
                        # KDL取込(EB)ボタン (Batch Index 1)
                        $kdlEbButton = New-Object System.Windows.Forms.Button
                        $kdlEbButton.Location = New-Object System.Drawing.Point($kdlEbX, $buttonY)
                        $kdlEbButton.Size = New-Object System.Drawing.Size(90, 30)
                        if ($script:editMode) { $kdlEbButton.Text = "参照" } else { $kdlEbButton.Text = "KDL取込(EB)" }
                        $kdlEbButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)
                        $kdlEbButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $kdlEbButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)
                        $kdlEbButton.FlatAppearance.BorderSize = 1
                        $kdlEbButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                        $kdlEbButton.Tag = @{ ProcessIndex = $i; BatchIndex = 1; Title = "KDL取込(EB)" }
                        $kdlEbButton.Add_Click({
                                $ctx = $this.Tag
                                $pIdx = $ctx.ProcessIndex
                                $bIdx = $ctx.BatchIndex
                                $title = $ctx.Title
                                if ($script:editMode) {
                                    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                    $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                    $fileDialog.Title = "$title 用バッチファイルを選択してください"

                                    # 初期ディレクトリの設定
                                    $initDir = ""
                                    
                                    # 既存の設定を確認
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $currentBatch = $procConf.BatchFiles[$bIdx]
                                            if ($currentBatch.Path) {
                                                $resolvedPath = Resolve-BatchPath -Path $currentBatch.Path
                                                if (Test-Path $resolvedPath) {
                                                    $initDir = Split-Path $resolvedPath -Parent
                                                }
                                            }
                                        }
                                    }
                                    
                                    # 既存がない、または無効な場合はGlobalLogPath/LogStoragePath
                                    if (-not $initDir -or -not (Test-Path $initDir)) {
                                        $logStoragePath = if ($script:globalLogPath) { $script:globalLogPath } else { "" }
                                        
                                        if (-not $logStoragePath) {
                                            $pageConfig = $script:pages[$script:currentPage]
                                            $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                            
                                            # メモリ上にない場合、JSONファイルから直接読み込む試み
                                            if (-not $logStoragePath -and $pageConfig.JsonPath) {
                                                $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
                                                    $pageConfig.JsonPath
                                                }
                                                else {
                                                    Join-Path $PSScriptRoot $pageConfig.JsonPath
                                                }
                                                
                                                if (Test-Path $jsonPath) {
                                                    try {
                                                        $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                                                        if ($pageJson.LogStoragePath) {
                                                            $logStoragePath = $pageJson.LogStoragePath
                                                            # メモリ上の設定も更新
                                                            $pageConfig.LogStoragePath = $logStoragePath
                                                        }
                                                    }
                                                    catch {}
                                                }
                                            }
                                        }

                                        if ($logStoragePath -and (Test-Path $initDir)) {
                                            $fileDialog.InitialDirectory = $initDir
                                        }
                                    }
                                    
                                    if ($initDir -and (Test-Path $initDir)) {
                                        $fileDialog.InitialDirectory = $initDir
                                    }

                                    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                        Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $fileDialog.FileName -BatchIndex $bIdx
                                        Update-ProcessControls
                                    }
                                    $fileDialog.Dispose()
                                }
                                else {
                                    # 引数の準備 (KdlSourcePath, KdlDestPath)
                                    $batchArgs = @()
                                    if ($script:processControls -and $pIdx -lt $script:processControls.Count) {
                                        $ctrls = $script:processControls[$pIdx]
                                        $src = if ($ctrls.KdlSourceTextBox) { $ctrls.KdlSourceTextBox.Text } else { "" }
                                        $dst = if ($ctrls.KdlDestTextBox) { $ctrls.KdlDestTextBox.Text } else { "" }
                                        if ($src -eq "パス") { $src = "" }
                                        if ($dst -eq "パス") { $dst = "" }
                                        $batchArgs = @($src, $dst)
                                    }

                                    # Get-BatchFilePathの代わり: インラインでパスを取得して解決
                                    $batchPath = $null
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $batch = $procConf.BatchFiles[$bIdx]
                                            $batchPath = Resolve-BatchPath -Path $batch.Path
                                        }
                                    }

                                    if ($batchPath) {
                                        Invoke-BatchFile -BatchPath $batchPath -DisplayName $title -ProcessIndex $pIdx -Arguments $batchArgs
                                        Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey "KdlEbButton_Executed"
                                        Update-ProcessControls
                                    }
                                }
                            })
                        # 編集モードOFF時：EnabledフラグをkdlEbボタンに反映
                        if (-not $script:editMode -and -not $isEnabled) {
                            $kdlEbButton.Enabled = $false
                        }
                        $script:processPanel.Controls.Add($kdlEbButton)
                        
                        # 直接取込ボタン (Batch Index 2) - Row 1用
                        $directImportButton = New-Object System.Windows.Forms.Button
                        $directImportButton.Location = New-Object System.Drawing.Point($directImportX_Row1, $buttonY)
                        $directImportButton.Size = New-Object System.Drawing.Size(90, 30)
                        if ($script:editMode) { $directImportButton.Text = "参照" } else { $directImportButton.Text = "直接取込" }
                        $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)
                        $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)
                        $directImportButton.FlatAppearance.BorderSize = 1
                        $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                        $directImportButton.Tag = @{ ProcessIndex = $i; BatchIndex = 2; Title = "直接取込" }
                        $directImportButton.Add_Click({
                                $ctx = $this.Tag
                                $pIdx = $ctx.ProcessIndex
                                $bIdx = $ctx.BatchIndex
                                $title = $ctx.Title
                                if ($script:editMode) {
                                    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                    $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                    $fileDialog.Title = "$title 用バッチファイルを選択してください"

                                    # 初期ディレクトリの設定
                                    $initDir = ""
                                    
                                    # 既存の設定を確認
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $currentBatch = $procConf.BatchFiles[$bIdx]
                                            if ($currentBatch.Path) {
                                                $resolvedPath = Resolve-BatchPath -Path $currentBatch.Path
                                                if (Test-Path $resolvedPath) {
                                                    $initDir = Split-Path $resolvedPath -Parent
                                                }
                                            }
                                        }
                                    }
                                    
                                    # 既存がない、または無効な場合はGlobalLogPath/LogStoragePath
                                    if (-not $initDir -or -not (Test-Path $initDir)) {
                                        $logStoragePath = if ($script:globalLogPath) { $script:globalLogPath } else { "" }
                                        
                                        if (-not $logStoragePath) {
                                            $pageConfig = $script:pages[$script:currentPage]
                                            $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                            
                                            # メモリ上にない場合、JSONファイルから直接読み込む試み
                                            if (-not $logStoragePath -and $pageConfig.JsonPath) {
                                                $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
                                                    $pageConfig.JsonPath
                                                }
                                                else {
                                                    Join-Path $PSScriptRoot $pageConfig.JsonPath
                                                }
                                                
                                                if (Test-Path $jsonPath) {
                                                    try {
                                                        $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                                                        if ($pageJson.LogStoragePath) {
                                                            $logStoragePath = $pageJson.LogStoragePath
                                                            # メモリ上の設定も更新
                                                            $pageConfig.LogStoragePath = $logStoragePath
                                                        }
                                                    }
                                                    catch {}
                                                }
                                            }
                                        }

                                        if ($logStoragePath -and (Test-Path $initDir)) {
                                            $fileDialog.InitialDirectory = $initDir
                                        }
                                    }
                                    
                                    if ($initDir -and (Test-Path $initDir)) {
                                        $fileDialog.InitialDirectory = $initDir
                                    }

                                    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                        Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $fileDialog.FileName -BatchIndex $bIdx
                                        Update-ProcessControls
                                    }
                                    $fileDialog.Dispose()
                                }
                                else {
                                    # 引数の準備 (KdlSourcePath, KdlDestPath)
                                    $batchArgs = @()
                                    if ($script:processControls -and $pIdx -lt $script:processControls.Count) {
                                        $ctrls = $script:processControls[$pIdx]
                                        $src = if ($ctrls.KdlSourceTextBox) { $ctrls.KdlSourceTextBox.Text } else { "" }
                                        $dst = if ($ctrls.KdlDestTextBox) { $ctrls.KdlDestTextBox.Text } else { "" }
                                        if ($src -eq "パス") { $src = "" }
                                        if ($dst -eq "パス") { $dst = "" }
                                        $batchArgs = @($src, $dst)
                                    }

                                    # Get-BatchFilePathの代わり: インラインでパスを取得して解決
                                    $batchPath = $null
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $pIdx -lt $currentProcesses.Count) {
                                        $procConf = $currentProcesses[$pIdx]
                                        if ($procConf.BatchFiles -and $procConf.BatchFiles.Count -gt $bIdx) {
                                            $batch = $procConf.BatchFiles[$bIdx]
                                            $batchPath = Resolve-BatchPath -Path $batch.Path
                                        }
                                    }

                                    if ($batchPath) {
                                        Invoke-BatchFile -BatchPath $batchPath -DisplayName $title -ProcessIndex $pIdx -Arguments $batchArgs
                                        Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey "DirectImportButton_Executed"
                                        Update-ProcessControls
                                    }
                                }
                            })
                        # 編集モードOFF時：EnabledフラグをdirectImportボタン（行1）に反映
                        if (-not $script:editMode -and -not $isEnabled) {
                            $directImportButton.Enabled = $false
                        }
                        $script:processPanel.Controls.Add($directImportButton)

                    }
                    else {
                        # --- 2行目 (Index 1) 標準レイアウト (変更なし) ---
                        
                        # KDL取込ボタン (Batch Index 0)
                        $kdlImportButton = New-Object System.Windows.Forms.Button
                        $kdlImportButton.Location = New-Object System.Drawing.Point($kdlImportX, $buttonY)
                        $kdlImportButton.Size = New-Object System.Drawing.Size(90, 30)
                        if ($script:editMode) { $kdlImportButton.Text = "参照" } else { $kdlImportButton.Text = "KDL取込" }
                        $kdlImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)
                        $kdlImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $kdlImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)
                        $kdlImportButton.FlatAppearance.BorderSize = 1
                        $kdlImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                        $kdlImportButton.Tag = $i
                        $kdlImportButton.Add_Click({
                                $clickedProcessIdx = $this.Tag
                                if ($script:editMode) {
                                    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                    $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                    $fileDialog.Title = "KDL取込用バッチファイルを選択してください"
                                    
                                    $pageConfig = $script:pages[$script:currentPage]
                                    $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                    if ($logStoragePath -and (Test-Path $logStoragePath)) { $fileDialog.InitialDirectory = $logStoragePath }
                                
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                        $processConfig = $currentProcesses[$clickedProcessIdx]
                                        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                            $currentBatch = $processConfig.BatchFiles[0]
                                            $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                            if (Test-Path $initialPath) {
                                                $fileDialog.InitialDirectory = Split-Path $initialPath
                                                $fileDialog.FileName = Split-Path $initialPath -Leaf
                                            }
                                        }
                                    }
                                
                                    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                        $selectedFile = $fileDialog.FileName
                                        if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 0) {
                                            Write-Log "KDL取込用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                            [System.Windows.Forms.MessageBox]::Show("KDL取込用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                            Update-ProcessControls
                                        }
                                    }
                                    $fileDialog.Dispose()
                                }
                                else {
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                        $processConfig = $currentProcesses[$clickedProcessIdx]
                                        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                            $batch = $processConfig.BatchFiles[0]
                                            $batchPath = Resolve-BatchPath -Path $batch.Path
                                            $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                            Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "KdlImportButton_Executed"
                                            Update-ProcessControls
                                        }
                                        else {
                                            Write-Log "KDL取込用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                            [System.Windows.Forms.MessageBox]::Show("KDL取込用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                        }
                                    }
                                }
                            })
                        # 編集モードOFF時：EnabledフラグをkdlImportボタン（行2）に反映
                        if (-not $script:editMode -and -not $isEnabled) {
                            $kdlImportButton.Enabled = $false
                        }
                        $script:processPanel.Controls.Add($kdlImportButton)
                        
                        # 直接取込ボタン (Batch Index 1) - Row 2用
                        $directImportButton = New-Object System.Windows.Forms.Button
                        $directImportButton.Location = New-Object System.Drawing.Point($directImportX, $buttonY)
                        $directImportButton.Size = New-Object System.Drawing.Size(90, 30)
                        if ($script:editMode) { $directImportButton.Text = "参照" } else { $directImportButton.Text = "直接取込" }
                        $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)
                        $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)
                        $directImportButton.FlatAppearance.BorderSize = 1
                        $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                        $directImportButton.Tag = $i
                        $directImportButton.Add_Click({
                                $clickedProcessIdx = $this.Tag
                                if ($script:editMode) {
                                    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                    $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                    $fileDialog.Title = "直接取込用バッチファイルを選択してください"
                                    
                                    $pageConfig = $script:pages[$script:currentPage]
                                    $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                    if ($logStoragePath -and (Test-Path $logStoragePath)) { $fileDialog.InitialDirectory = $logStoragePath }
                                    
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                        $processConfig = $currentProcesses[$clickedProcessIdx]
                                        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                            $currentBatch = $processConfig.BatchFiles[1]
                                            $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                            if (Test-Path $initialPath) {
                                                $fileDialog.InitialDirectory = Split-Path $initialPath
                                                $fileDialog.FileName = Split-Path $initialPath -Leaf
                                            }
                                        }
                                        elseif ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                            $currentBatch = $processConfig.BatchFiles[0]
                                            $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                            if (Test-Path $initialPath) { $fileDialog.InitialDirectory = Split-Path $initialPath }
                                        }
                                    }
                                    
                                    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                        $selectedFile = $fileDialog.FileName
                                        Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 1
                                        Write-Log "直接取込用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("直接取込用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                        Update-ProcessControls
                                    }
                                    $fileDialog.Dispose()
                                }
                                else {
                                    $currentProcesses = Get-CurrentPageProcesses
                                    if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                        $processConfig = $currentProcesses[$clickedProcessIdx]
                                        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                            $batch = $processConfig.BatchFiles[1]
                                            $batchPath = Resolve-BatchPath -Path $batch.Path
                                            $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                            Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "DirectImportButton_Executed"
                                            Update-ProcessControls
                                        }
                                        else {
                                            Write-Log "直接取込用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                            [System.Windows.Forms.MessageBox]::Show("直接取込用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                        }
                                    }
                                }
                            })
                        # 編集モードOFF時：EnabledフラグをdirectImportボタン（行2）に反映
                        if (-not $script:editMode -and -not $isEnabled) {
                            $directImportButton.Enabled = $false
                        }
                        $script:processPanel.Controls.Add($directImportButton)
                    }
                    
                    $maintButton1 = $null
                    $maintButton2 = $null
                    $afterImportButton = $null # Initialize for scope

                    if ($i -eq 0) {
                        # --- 1行目 (Index 0) ---
                        # 取込後EA (Batch 3),取込後EB (Batch 4), メンテ(Batch 5)
                        
                        # 取込後EA (X=540)
                        $afterImportButton = Create-AfterImportButton -x 540 -batchIndex 3 -text "取込後EA"
                        
                        # 取込後EB (X=630)
                        $maintButton1 = Create-MaintButton -x 630 -text "取込後EB" -batchIndex 4 -title "取込後EB" -componentKey "MaintButton1_Executed"
                        
                        # メンテ (X=720)
                        $maintButton2 = Create-MaintButton -x 720 -text "メンテ" -batchIndex 5 -title "メンテ" -componentKey "MaintButton2_Executed"
                        
                        # ログ確認 (Y+35)
                        # Log1 @ 240, Log2 @ 440 (Direct Import X)
                        $logY = $buttonY + 35
                        $log1X = 240
                        $log2X = 440
                    }
                    else {
                        # --- 2行目 (Index 1) & 3行目 (Index 2) ---
                        # 取込後 (Batch 2)
                        
                        # 取込後 (X=440)
                        $afterImportButton = Create-AfterImportButton -x 440 -batchIndex 2 -text "取込後"
                        
                        # メンテ (X=530) - 3行目のみ
                        if ($i -eq 2) {
                            $maintButton1 = Create-MaintButton -x 530 -text "メンテ" -batchIndex 3 -title "メンテ"
                        }
                        
                        # ログ確認 (Y+35)
                        # Log1 @ 240, Log2 @ 340 (Direct Import X)
                        $logY = $buttonY + 35
                        $log1X = 240
                        $log2X = 340
                    }
                    
                    # ログ確認ボタン1 (KDL列)
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point($log1X, $logY)
                    $logButton.Size = New-Object System.Drawing.Size(90, 30)
                    if ($script:editMode) { $logButton.Text = "参照" } else { $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" } }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton.Tag = $i
                    $logButton.Add_Click({ $clickedProcessIdx = $this.Tag; Show-ProcessLog -ProcessIndex $clickedProcessIdx })
                    $script:processPanel.Controls.Add($logButton)
                    
                    # ログ確認ボタン2 (Direct列)
                    $logButton2 = New-Object System.Windows.Forms.Button
                    $logButton2.Location = New-Object System.Drawing.Point($log2X, $logY)
                    $logButton2.Size = New-Object System.Drawing.Size(90, 30)
                    if ($script:editMode) { $logButton2.Text = "参照" } else { $logButton2.Text = "ログ確認" }
                    $logButton2.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)
                    $logButton2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton2.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)
                    $logButton2.FlatAppearance.BorderSize = 1
                    $logButton2.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton2.Tag = $i
                    $logButton2.Add_Click({ $clickedProcessIdx = $this.Tag; Show-ProcessLog -ProcessIndex $clickedProcessIdx -LogIndex 2 })
                    $script:processPanel.Controls.Add($logButton2)

                    # 4ページ目用のコントロール情報を保存
                    $script:processControls += @{
                        CheckBox            = $null
                        NameTextBox         = $nameTextBox
                        KdlSourceTextBox    = $kdlSourceTextBox
                        KdlSourceMoveButton = $null
                        KdlDestTextBox      = $kdlDestTextBox
                        KdlDestMoveButton   = $kdlDestMoveButton
                        V1CsvDestTextBox    = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        KdlImportButton     = $kdlImportButton
                        KdlKdbButton        = $kdlKdbButton
                        KdlEbButton         = $kdlEbButton
                        DirectImportButton  = $directImportButton
                        AfterImportButton   = $afterImportButton
                        MaintButton1        = $maintButton1
                        MaintButton2        = $maintButton2
                        LogButton           = $logButton
                        LogButton2          = $logButton2
                    }
                }
                elseif ($i -eq 2) {
                    # 3行目 (Index 2): V1抽出CSV格納先 + 特定ボタン (取込後、メンテ、ログ)
                    
                    # チェックボックスは表示しない (固定行扱い)
                    
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(175, [int]($y - 20))
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    
                    # V1抽出CSV格納先パス入力
                    $v1CsvDestTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvDestTextBox.Location = New-Object System.Drawing.Point(175, $y)
                    $v1CsvDestTextBox.Size = New-Object System.Drawing.Size(200, 30)
                    $v1CsvDestTextBox.Text = "パス"
                    $v1CsvDestTextBox.ReadOnly = $true
                    $v1CsvDestTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvDestTextBox.Tag = $i
                    $v1CsvDestTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $this.Text -Description "V1抽出CSV格納先フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessV1CsvDestPath -ProcessIndex $clickedProcessIdx -V1CsvDestPath $selectedPath
                                    Write-Log "V1抽出CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    
                    # 初期値設定
                    $v1CsvDestPathValue = "パス"
                    if ($processConfig.V1CsvDestPath -and $processConfig.V1CsvDestPath -ne "" -and $processConfig.V1CsvDestPath -ne "パス") {
                        try {
                            $v1CsvDestPathValue = $processConfig.V1CsvDestPath
                            if (-not [System.IO.Path]::IsPathRooted($v1CsvDestPathValue)) {
                                $basePath = Get-CommonBasePath
                                $v1CsvDestPathValue = Join-Path $basePath $v1CsvDestPathValue
                            }
                            $v1CsvDestPathValue = [System.IO.Path]::GetFullPath($v1CsvDestPathValue)
                        }
                        catch {
                            Write-Log "パスの解決に失敗しました (Process: $i): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $v1CsvDestTextBox.Text = $v1CsvDestPathValue
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # 移動設定ボタン
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(385, $y)
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) {
                        $v1CsvDestMoveButton.Text = "移動設定"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)
                    }
                    else {
                        $v1CsvDestMoveButton.Text = "移動"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)
                    }
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $true
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            $v1CsvDestPath = ""
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup.NameTextBox) { $currentProcessName = $ctrlGroup.NameTextBox.Text }
                                if ($ctrlGroup.V1CsvDestTextBox) { $v1CsvDestPath = $ctrlGroup.V1CsvDestTextBox.Text }
                            }
                            $v1CsvSourcePath = if ($script:v1CsvSourceTextBox) { $script:v1CsvSourceTextBox.Text } else { "" }
                            
                            if ($script:editMode) {
                                Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName
                            }
                            else {
                                Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $v1CsvSourcePath -V1CsvDestinationPath $v1CsvDestPath
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    # ボタン行（直接取込、取込後、メンテ、ログ確認）
                    $buttonY = $y
                    
                    # 直接取込ボタン（Batch Index 1）
                    $directImportButton = New-Object System.Windows.Forms.Button
                    $directImportButton.Location = New-Object System.Drawing.Point(450, $buttonY)
                    $directImportButton.Size = New-Object System.Drawing.Size(90, 30)

                    if ($script:editMode) {
                        $directImportButton.Text = "参照"
                    }
                    else {
                        $directImportButton.Text = "直接取込"
                    }
                    $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)  # #ffe6cc
                    $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)  # #d79b00
                    $directImportButton.FlatAppearance.BorderSize = 1
                    $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $directImportButton.Tag = $i
                    $directImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # バッチファイル設定 (Index 1)
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "直接取込用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $fileDialog.FileName -BatchIndex 1
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 実行 (Batch Index 1)
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $procConf = $currentProcesses[$clickedProcessIdx]
                                    if ($procConf.BatchFiles.Count -gt 1) {
                                        $batch = $procConf.BatchFiles[1]
                                        $path = Resolve-BatchPath -Path $batch.Path
                                        Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "DirectImportButton_Executed"
                                        Update-ProcessControls
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($directImportButton)

                    # 取込後ボタン (Batch Index 2)
                    $afterImportButton = New-Object System.Windows.Forms.Button
                    $afterImportButton.Location = New-Object System.Drawing.Point(540, $y)
                    $afterImportButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $afterImportButton.Text = "参照"
                    }
                    else {
                        $afterImportButton.Text = "取込後"
                    }
                    $afterImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)
                    $afterImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $afterImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)
                    $afterImportButton.FlatAppearance.BorderSize = 1
                    $afterImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $afterImportButton.Tag = $i
                    $afterImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # バッチファイル設定 (Index 2)
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "取込後用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $fileDialog.FileName -BatchIndex 2
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 実行 (Batch Index 2)
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $procConf = $currentProcesses[$clickedProcessIdx]
                                    if ($procConf.BatchFiles.Count -gt 2) {
                                        $batch = $procConf.BatchFiles[2]
                                        $path = Resolve-BatchPath -Path $batch.Path
                                        Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "AfterImportButton_Executed"
                                        Update-ProcessControls
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($afterImportButton)
                    
                    $script:processPanel.Controls.Add($afterImportButton)
                    
                    # メンテボタン (Batch Index 3)
                    $maintButton = New-Object System.Windows.Forms.Button
                    $maintButton.Location = New-Object System.Drawing.Point(630, $y)
                    $maintButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) { $maintButton.Text = "参照" } else { $maintButton.Text = "メンテ" }
                    $maintButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153) # Orange
                    $maintButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $maintButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)
                    $maintButton.FlatAppearance.BorderSize = 1
                    $maintButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $maintButton.Tag = @{ ProcessIndex = $i; BatchIndex = 3; Title = "メンテ" }
                    $maintButton.Add_Click({
                            $ctx = $this.Tag
                            $pIdx = $ctx.ProcessIndex
                            $bIdx = $ctx.BatchIndex
                            $title = $ctx.Title
                            
                            if ($script:editMode) {
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "$title 用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $fileDialog.FileName -BatchIndex $bIdx
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                $currentProcesses = Get-CurrentPageProcesses
                                $procConf = $currentProcesses[$pIdx]
                                if ($procConf.BatchFiles.Count -gt $bIdx) {
                                    $batch = $procConf.BatchFiles[$bIdx]
                                    $path = Resolve-BatchPath -Path $batch.Path
                                    Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $pIdx
                                    Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey "MaintButton_Executed"
                                    Update-ProcessControls
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($maintButton)
                    
                    # ログ確認ボタン
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point(720, $y)
                    $logButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) { $logButton.Text = "参照" } else { $logButton.Text = "ログ確認" }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton.Tag = $i
                    $logButton.Add_Click({ Show-ProcessLog -ProcessIndex $this.Tag })
                    $script:processPanel.Controls.Add($logButton)
                    
                    $script:processControls += @{
                        NameTextBox         = $nameTextBox
                        V1CsvDestTextBox    = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        DirectImportButton  = $directImportButton
                        AfterImportButton   = $afterImportButton
                        MaintButton         = $maintButton
                        KdlImportButton     = $null
                        KdlKdbButton        = $null
                        KdlEbButton         = $null
                        MaintButton1        = $null
                        MaintButton2        = $null
                        LogButton           = $logButton
                    }
                }
                else {
                    # 4行目以降（Index 3以上）：V1抽出CSV格納先のみ
                    # 有効/無効切り替え用チェックボックス（Page 4 Index 3+ 用）
                    $calcX = $x - 25
                    $calcY = $y + 5
                    $enableCheckBox = New-Object System.Windows.Forms.CheckBox
                    $enableCheckBox.Location = New-Object System.Drawing.Point($calcX, $calcY)
                    $enableCheckBox.Size = New-Object System.Drawing.Size(20, 20)
                    $enableCheckBox.Checked = $isEnabled
                    $enableCheckBox.Visible = $script:editMode
                    $enableCheckBox.Tag = $i
                    $enableCheckBox.Add_Click({
                            $idx = $this.Tag
                            $enabled = $this.Checked
                            Save-ProcessEnabled -ProcessIndex $idx -Enabled $enabled
                            Update-ProcessControls
                        })
                    $script:processPanel.Controls.Add($enableCheckBox)
                    
                    # 行削除用チェックボックス（編集モードON時のみ表示）
                    $checkBox = New-Object System.Windows.Forms.CheckBox
                    $checkBox.Location = New-Object System.Drawing.Point($calcX, [int]($calcY + 25))
                    $checkBox.Size = New-Object System.Drawing.Size(20, 20)
                    $checkBox.Visible = $script:editMode
                    $checkBox.Tag = $i
                    $script:processPanel.Controls.Add($checkBox)
                    
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(175, [int]($y - 20))
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    
                    # V1抽出CSV格納先パス入力
                    $v1CsvDestTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvDestTextBox.Location = New-Object System.Drawing.Point(175, $y)
                    $v1CsvDestTextBox.Size = New-Object System.Drawing.Size(200, 30)
                    $v1CsvDestTextBox.Text = "パス"
                    $v1CsvDestTextBox.ReadOnly = $true
                    $v1CsvDestTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvDestTextBox.Tag = $i  # プロセスインデックスをTagに保存（3行目以降）
                    $v1CsvDestTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $this.Text -Description "V1抽出CSV格納先フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    # 各プロセスのV1CsvDestPathをpage4.jsonに保存
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessV1CsvDestPath -ProcessIndex $clickedProcessIdx -V1CsvDestPath $selectedPath
                                    Write-Log "V1抽出CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    # V1抽出CSV格納先の初期値を設定（3行目以降）
                    $v1CsvDestPathValue = "パス"
                    if ($processConfig.V1CsvDestPath -and $processConfig.V1CsvDestPath -ne "" -and $processConfig.V1CsvDestPath -ne "パス") {
                        try {
                            $v1CsvDestPathValue = $processConfig.V1CsvDestPath
                            # 相対パスの場合は絶対パスに変換
                            if (-not [System.IO.Path]::IsPathRooted($v1CsvDestPathValue)) {
                                $basePath = Get-CommonBasePath
                                $v1CsvDestPathValue = Join-Path $basePath $v1CsvDestPathValue
                            }
                            $v1CsvDestPathValue = [System.IO.Path]::GetFullPath($v1CsvDestPathValue)
                        }
                        catch {
                            # エラー時はデフォルト値を使用
                            Write-Log "パスの解決に失敗しました (Process: $i, Type: V1CsvDestPath): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $v1CsvDestTextBox.Text = $v1CsvDestPathValue
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # V1抽出CSV格納先の移動設定ボタン（編集モードON時は水色、OFF時は紺色）
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(385, $y)
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) {
                        $v1CsvDestMoveButton.Text = "移動設定"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    }
                    else {
                        $v1CsvDestMoveButton.Text = "移動"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                    }
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $true  # 常に表示
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            $v1CsvSourcePath = ""
                            $v1CsvDestPath = ""
                        
                            # プロセス名とパスを取得
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
                                    $currentProcessName = $ctrlGroup.NameTextBox.Text
                                }
                                if ($ctrlGroup -and $ctrlGroup.V1CsvDestTextBox) {
                                    $v1CsvDestPath = $ctrlGroup.V1CsvDestTextBox.Text
                                }
                            }
                        
                            # V1抽出CSV格納元を取得
                            if ($script:v1CsvSourceTextBox) {
                                $v1CsvSourcePath = $script:v1CsvSourceTextBox.Text
                            }
                        
                            # 編集モードと非編集モードで動作を分岐
                            if ($script:editMode) {
                                # 編集モード：移動設定ダイアログを表示
                                Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName
                            }
                            else {
                                # 非編集モード：ファイル移動を実行
                                Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $v1CsvSourcePath -V1CsvDestinationPath $v1CsvDestPath
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    # ボタン行（直接取込、取込後、ログ確認）
                    $buttonY = $y
                    
                    # 直接取込ボタン（オレンジ）
                    $directImportButton = New-Object System.Windows.Forms.Button
                    $directImportButton.Location = New-Object System.Drawing.Point(450, $buttonY)
                    $directImportButton.Size = New-Object System.Drawing.Size(90, 30)

                    if ($script:editMode) {
                        $directImportButton.Text = "参照"
                    }
                    else {
                        $directImportButton.Text = "直接取込"
                    }
                    $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)  # #ffe6cc
                    $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)  # #d79b00
                    $directImportButton.FlatAppearance.BorderSize = 1
                    $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $directImportButton.Tag = $i  # プロセスインデックスをTagに保存
                    $directImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # 編集モードON：ファイル選択ダイアログでバッチファイルのパスをJSONに保存
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "直接取込用バッチファイルを選択してください"
                                
                                # LogStoragePathを初期ディレクトリに設定
                                $pageConfig = $script:pages[$script:currentPage]
                                $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                if ($logStoragePath -and (Test-Path $logStoragePath)) {
                                    $fileDialog.InitialDirectory = $logStoragePath
                                }
                            
                                # 現在のバッチファイルパスを初期値として設定（BatchIndex = 1）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                        $currentBatch = $processConfig.BatchFiles[1]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                                        }
                                    }
                                    elseif ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        # BatchFiles[1]が存在しない場合は、BatchFiles[0]を初期値として使用
                                        $currentBatch = $processConfig.BatchFiles[0]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                        }
                                    }
                                }
                            
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $selectedFile = $fileDialog.FileName
                                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 1) {
                                        Write-Log "直接取込用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("直接取込用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                        # コントロールを更新して新しい設定を反映
                                        Update-ProcessControls
                                    }
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 編集モードOFF：JSONに設定されたバッチファイルを実行（BatchIndex = 1）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 1) {
                                        $batch = $processConfig.BatchFiles[1]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $batchPath = Resolve-BatchPath -Path $batch.Path
                                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "DirectImportButton_Executed"
                                        Update-ProcessControls
                                    }
                                    else {
                                        Write-Log "直接取込用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("直接取込用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                    }
                                }
                            }
                        })
                    # 編集モードOFF時：EnabledフラグをdirectImportボタン（行3以降）に反映
                    if (-not $script:editMode -and -not $isEnabled) {
                        $directImportButton.Enabled = $false
                    }
                    $script:processPanel.Controls.Add($directImportButton)
                    
                    # 取込後ボタン（オレンジ）
                    $afterImportButton = New-Object System.Windows.Forms.Button
                    $afterImportButton.Location = New-Object System.Drawing.Point(540, $buttonY)
                    $afterImportButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $afterImportButton.Text = "参照"
                    }
                    else {
                        $afterImportButton.Text = "取込後"
                    }
                    $afterImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                    $afterImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $afterImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                    $afterImportButton.FlatAppearance.BorderSize = 1
                    $afterImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $afterImportButton.Tag = $i  # プロセスインデックスをTagに保存
                    $afterImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # 編集モードON：ファイル選択ダイアログでバッチファイルのパスをJSONに保存
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "取込後用バッチファイルを選択してください"
                                
                                # LogStoragePathを初期ディレクトリに設定
                                $pageConfig = $script:pages[$script:currentPage]
                                $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
                                if ($logStoragePath -and (Test-Path $logStoragePath)) {
                                    $fileDialog.InitialDirectory = $logStoragePath
                                }
                            
                                # 現在のバッチファイルパスを初期値として設定（BatchIndex = 2）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 2) {
                                        $currentBatch = $processConfig.BatchFiles[2]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                            $fileDialog.FileName = Split-Path $initialPath -Leaf
                                        }
                                    }
                                    elseif ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                        # BatchFiles[2]が存在しない場合は、BatchFiles[0]を初期値として使用
                                        $currentBatch = $processConfig.BatchFiles[0]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $initialPath = Resolve-BatchPath -Path $currentBatch.Path
                                        if (Test-Path $initialPath) {
                                            $fileDialog.InitialDirectory = Split-Path $initialPath
                                        }
                                    }
                                }
                            
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $selectedFile = $fileDialog.FileName
                                    if (Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 2) {
                                        Write-Log "取込後用バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("取込後用バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    
                                        # コントロールを更新して新しい設定を反映
                                        Update-ProcessControls
                                    }
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 編集モードOFF：JSONに設定されたバッチファイルを実行（BatchIndex = 2）
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $processConfig = $currentProcesses[$clickedProcessIdx]
                                    if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 2) {
                                        $batch = $processConfig.BatchFiles[2]
                                        # Resolve-BatchPathを使用してパスを解決
                                        $batchPath = Resolve-BatchPath -Path $batch.Path
                                        $result = Invoke-BatchFile -BatchPath $batchPath -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                        Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "AfterImportButton_Executed"
                                        Update-ProcessControls
                                    }
                                    else {
                                        Write-Log "取込後用バッチファイルが設定されていません" "ERROR" $clickedProcessIdx
                                        [System.Windows.Forms.MessageBox]::Show("取込後用バッチファイルが設定されていません。`n編集モードでバッチファイルを設定してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                    }
                                }
                            }
                        })
                    # 編集モードOFF時：EnabledフラグをafterImportボタン（行3以降）に反映
                    if (-not $script:editMode -and -not $isEnabled) {
                        $afterImportButton.Enabled = $false
                    }
                    $script:processPanel.Controls.Add($afterImportButton)
                    
                    # ログ確認ボタン（緑）
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point(630, $buttonY)
                    $logButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $logButton.Text = "参照"
                    }
                    else {
                        $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                    }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton.Tag = $i
                    $logButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            Show-ProcessLog -ProcessIndex $clickedProcessIdx
                        })
                    $script:processPanel.Controls.Add($logButton)
                    
                    
                    # 4ページ目用のコントロール情報を保存（3行目以降）
                    $script:processControls += @{
                        CheckBox            = $checkBox
                        NameTextBox         = $nameTextBox
                        V1CsvDestTextBox    = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        DirectImportButton  = $directImportButton
                        AfterImportButton   = $afterImportButton
                        LogButton           = $logButton
                    }
                }
            }
            else {
                # 3行目以降（Index 2以上、Page 4の場合）または5ページ目以降
                
                # 座標計算
                if ($isPage4) {
                    # 4ページ目：1列レイアウト
                    # $yはループの先頭で計算済み ($y = 10 + $i * 170)
                    $x = 55
                    # $y = $y # 既に計算されている値をそのまま使用
                }
                else {
                    # 5ページ目以降：2列レイアウト
                    $row = [Math]::Floor($i / 2)
                    $col = $i % 2
                    $x = [int](10 + $col * 440)
                    $y = [int](10 + $row * 60)
                }
                
                # チェックボックス（編集モードON時のみ表示）
                $checkBox = New-Object System.Windows.Forms.CheckBox
                
                if ($isPage4) {
                    # 4ページ目：Index 3以降のみ表示
                    $checkBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 5))
                    if ($i -lt 3) {
                        $checkBox.Visible = $false
                    }
                    else {
                        $checkBox.Visible = $script:editMode
                    }
                }
                else {
                    # 5ページ目以降
                    $checkBox.Location = New-Object System.Drawing.Point([int]($x - 25), [int]($y + 10))
                    $checkBox.Visible = $script:editMode
                }
                
                $checkBox.Size = New-Object System.Drawing.Size(20, 20)
                $script:processPanel.Controls.Add($checkBox)
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(140, 40)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = -not $script:editMode
                $nameTextBox.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 255)
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                $nameTextBox.Multiline = $true
                $nameTextBox.Height = 40
                $nameTextBox.Tag = $i
                $nameTextBox.Add_Leave({
                        if ($script:editMode) {
                            $processIdx = $this.Tag
                            $newName = $this.Text
                            Save-ProcessName -ProcessIndex $processIdx -ProcessName $newName
                        }
                    })
                $script:processPanel.Controls.Add($nameTextBox)
                
                # V1抽出CSV格納先（4ページ目のみ表示）
                if ($isPage4) {
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(175, [int]($y - 20))
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    
                    # V1抽出CSV格納先パス入力
                    $v1CsvDestTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvDestTextBox.Location = New-Object System.Drawing.Point(175, $y)
                    $v1CsvDestTextBox.Size = New-Object System.Drawing.Size(200, 30)
                    $v1CsvDestTextBox.Text = "パス"
                    $v1CsvDestTextBox.ReadOnly = $true
                    $v1CsvDestTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvDestTextBox.Tag = $i
                    $v1CsvDestTextBox.Add_Click({
                            if ($script:editMode) {
                                $selectedPath = Show-FolderBrowser -InitialDirectory $this.Text -Description "V1抽出CSV格納先フォルダを選択してください"
                                if ($selectedPath) {
                                    $this.Text = $selectedPath
                                    $clickedProcessIdx = $this.Tag
                                    Save-ProcessV1CsvDestPath -ProcessIndex $clickedProcessIdx -V1CsvDestPath $selectedPath
                                    Write-Log "V1抽出CSV格納先を設定しました: $selectedPath" "INFO" $clickedProcessIdx
                                }
                            }
                            else {
                                $path = $this.Text
                                if (-not [string]::IsNullOrWhiteSpace($path) -and $path -ne "パス") {
                                    if (-not [System.IO.Path]::IsPathRooted($path)) {
                                        $path = Join-Path $PSScriptRoot $path
                                    }
                                    if (Test-Path $path) {
                                        Open-PathInExplorer -Path $path
                                    }
                                    else {
                                        Write-Log "パスが存在しません: $path" "WARN"
                                    }
                                }
                            }
                        })
                    
                    # 初期値設定
                    $v1CsvDestPathValue = "パス"
                    if ($processConfig.V1CsvDestPath -and $processConfig.V1CsvDestPath -ne "" -and $processConfig.V1CsvDestPath -ne "パス") {
                        try {
                            $v1CsvDestPathValue = $processConfig.V1CsvDestPath
                            if (-not [System.IO.Path]::IsPathRooted($v1CsvDestPathValue)) {
                                $basePath = Get-CommonBasePath
                                $v1CsvDestPathValue = Join-Path $basePath $v1CsvDestPathValue
                            }
                            $v1CsvDestPathValue = [System.IO.Path]::GetFullPath($v1CsvDestPathValue)
                        }
                        catch {
                            Write-Log "パスの解決に失敗しました (Process: $i): $($_.Exception.Message)" "WARN"
                        }
                    }
                    $v1CsvDestTextBox.Text = $v1CsvDestPathValue
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # 移動設定ボタン
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(385, $y)
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) {
                        $v1CsvDestMoveButton.Text = "移動設定"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)
                    }
                    else {
                        $v1CsvDestMoveButton.Text = "移動"
                        $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
                        $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)
                    }
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $true
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            $v1CsvDestPath = ""
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup.NameTextBox) { $currentProcessName = $ctrlGroup.NameTextBox.Text }
                                if ($ctrlGroup.V1CsvDestTextBox) { $v1CsvDestPath = $ctrlGroup.V1CsvDestTextBox.Text }
                            }
                            $v1CsvSourcePath = if ($script:v1CsvSourceTextBox) { $script:v1CsvSourceTextBox.Text } else { "" }
                            
                            if ($script:editMode) {
                                Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName
                            }
                            else {
                                Invoke-FileMoveOperation -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName -V1CsvSourcePath $v1CsvSourcePath -V1CsvDestinationPath $v1CsvDestPath
                            }
                        })
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    
                    # ボタン配置 (Row 4+, Index 3+)
                    $buttonY = $y
                    
                    # 直接取込ボタン (Batch Index 1)
                    $directImportButton = New-Object System.Windows.Forms.Button
                    $directImportButton.Location = New-Object System.Drawing.Point(450, $buttonY)
                    $directImportButton.Size = New-Object System.Drawing.Size(90, 30)
                    if ($script:editMode) {
                        $directImportButton.Text = "参照"
                    }
                    else {
                        $directImportButton.Text = "直接取込"
                    }
                    $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)  # #ffe6cc
                    $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)  # #d79b00
                    $directImportButton.FlatAppearance.BorderSize = 1
                    $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $directImportButton.Tag = $i
                    $directImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # バッチファイル設定 (Index 1)
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "直接取込用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $fileDialog.FileName -BatchIndex 1
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 実行 (Batch Index 1)
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $procConf = $currentProcesses[$clickedProcessIdx]
                                    if ($procConf.BatchFiles.Count -gt 1) {
                                        $batch = $procConf.BatchFiles[1]
                                        $path = Resolve-BatchPath -Path $batch.Path
                                        Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($directImportButton)
                    
                    # 取込後ボタン (Batch Index 2)
                    $afterImportButton = New-Object System.Windows.Forms.Button
                    $afterImportButton.Location = New-Object System.Drawing.Point(540, $buttonY)
                    $afterImportButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $afterImportButton.Text = "参照"
                    }
                    else {
                        $afterImportButton.Text = "取込後"
                    }
                    $afterImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)
                    $afterImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $afterImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)
                    $afterImportButton.FlatAppearance.BorderSize = 1
                    $afterImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $afterImportButton.Tag = $i
                    $afterImportButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            if ($script:editMode) {
                                # バッチファイル設定 (Index 2)
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "取込後用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $fileDialog.FileName -BatchIndex 2
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                # 実行 (Batch Index 2)
                                $currentProcesses = Get-CurrentPageProcesses
                                if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                    $procConf = $currentProcesses[$clickedProcessIdx]
                                    if ($procConf.BatchFiles.Count -gt 2) {
                                        $batch = $procConf.BatchFiles[2]
                                        $path = Resolve-BatchPath -Path $batch.Path
                                        Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $clickedProcessIdx
                                    }
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($afterImportButton)

                    # メンテボタン (Batch Index 3)
                    $maintButton = New-Object System.Windows.Forms.Button
                    $maintButton.Location = New-Object System.Drawing.Point(630, $buttonY)
                    $maintButton.Size = New-Object System.Drawing.Size(60, 30)
                    if ($script:editMode) { $maintButton.Text = "参照" } else { $maintButton.Text = "メンテ" }
                    $maintButton.BackColor = [System.Drawing.Color]::FromArgb(255, 192, 203) # Pink
                    $maintButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $maintButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 100, 100)
                    $maintButton.FlatAppearance.BorderSize = 1
                    $maintButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $maintButton.Tag = @{ ProcessIndex = $i; BatchIndex = 3; Title = "メンテ" }
                    $maintButton.Add_Click({
                            $ctx = $this.Tag
                            $pIdx = $ctx.ProcessIndex
                            $bIdx = $ctx.BatchIndex
                            $title = $ctx.Title
                            
                            if ($script:editMode) {
                                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                                $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
                                $fileDialog.Title = "$title 用バッチファイルを選択してください"
                                if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    Save-BatchFilePath -ProcessIndex $pIdx -BatchFilePath $fileDialog.FileName -BatchIndex $bIdx
                                    Update-ProcessControls
                                }
                                $fileDialog.Dispose()
                            }
                            else {
                                $currentProcesses = Get-CurrentPageProcesses
                                $procConf = $currentProcesses[$pIdx]
                                if ($procConf.BatchFiles.Count -gt $bIdx) {
                                    $batch = $procConf.BatchFiles[$bIdx]
                                    $path = Resolve-BatchPath -Path $batch.Path
                                    Invoke-BatchFile -BatchPath $path -DisplayName $batch.Name -ProcessIndex $pIdx
                                    Save-ProcessComponentExecuted -ProcessIndex $pIdx -ComponentKey "MaintButton_Executed"
                                    Update-ProcessControls
                                }
                            }
                        })
                    $script:processPanel.Controls.Add($maintButton)

                    # ログ確認ボタン
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point(720, $buttonY)
                    $logButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) { $logButton.Text = "参照" } else { $logButton.Text = "ログ確認" }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton.Tag = $i
                    $logButton.Add_Click({ Show-ProcessLog -ProcessIndex $this.Tag })
                    $script:processPanel.Controls.Add($logButton)
                    
                    $script:processControls += @{
                        CheckBox            = $checkBox
                        NameTextBox         = $nameTextBox
                        V1CsvDestTextBox    = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        DirectImportButton  = $directImportButton
                        AfterImportButton   = $afterImportButton
                        MaintButton         = $maintButton
                        LogButton           = $logButton
                        ExecuteButton       = $null
                    }
                }
                else {
                    # 5ページ目以降 (Existing logic)
                    # ファイル移動設定ボタン（水色）
                    $fileMoveButton = New-Object System.Windows.Forms.Button
                    $fileMoveX = [int]($x + 150)
                    $fileMoveButton.Location = New-Object System.Drawing.Point($fileMoveX, $y)
                    $fileMoveButton.Size = New-Object System.Drawing.Size(80, 40)
                    $fileMoveButton.Text = "移動設定"
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
                    $fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $fileMoveButton.FlatAppearance.BorderSize = 1
                    $fileMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $fileMoveButton.Visible = $script:editMode
                    $fileMoveButton.Tag = $i
                    $fileMoveButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            $currentProcessName = ""
                            if ($script:processControls -and $clickedProcessIdx -lt $script:processControls.Count) {
                                $ctrlGroup = $script:processControls[$clickedProcessIdx]
                                if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
                                    $currentProcessName = $ctrlGroup.NameTextBox.Text
                                }
                            }
                            Show-FileMoveSettingsDialog -ProcessIndex $clickedProcessIdx -ProcessName $currentProcessName
                        })
                    $script:processPanel.Controls.Add($fileMoveButton)
                    
                    # 実行ボタン（オレンジ）
                    $executeButton = New-Object System.Windows.Forms.Button
                    $executeX = [int]($x + 240)
                    $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
                    $executeButton.Size = New-Object System.Drawing.Size(80, 40)
                    if ($script:editMode) {
                        $executeButton.Text = "参照"
                    }
                    else {
                        $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                    }
                    $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
                    $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $executeButton.FlatAppearance.BorderSize = 1
                    $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $executeButton.Tag = $i  # プロセスインデックスをTagに保存
                    $executeButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            Start-ProcessFlow -ProcessIndex $clickedProcessIdx
                            # Start-ProcessFlow内で個別ボタンの処理を行うため、ここでは全体非活性化を削除
                            # Save-ProcessComponentExecuted -ProcessIndex $clickedProcessIdx -ComponentKey "FileMoveButton_Executed"
                            Update-ProcessControls
                        })
                    $script:processPanel.Controls.Add($executeButton)
                    
                    # ログ確認ボタン（緑）
                    $logButton = New-Object System.Windows.Forms.Button
                    $logX = [int]($x + 330)
                    $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                    $logButton.Size = New-Object System.Drawing.Size(80, 40)
                    if ($script:editMode) {
                        $logButton.Text = "参照"
                    }
                    else {
                        $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                    }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $logButton.Tag = $i  # プロセスインデックスをTagに保存
                    $logButton.Add_Click({
                            $clickedProcessIdx = $this.Tag
                            Show-ProcessLog -ProcessIndex $clickedProcessIdx
                        })
                    $script:processPanel.Controls.Add($logButton)
                    
                    # 5ページ目以降用のコントロール情報を保存
                    $script:processControls += @{
                        CheckBox       = $checkBox
                        NameTextBox    = $nameTextBox
                        FileMoveButton = $fileMoveButton
                        ExecuteButton  = $executeButton
                        LogButton      = $logButton
                    }
                }
            }
        }
    }
    
    # ページ情報の更新
    $script:pageLabel.Text = "ページ $($script:currentPage + 1) / $totalPages"
    
    # 行追加・削除ボタンの表示/非表示を編集モードに応じて更新
    if ($script:addRowButton) {
        $script:addRowButton.Visible = $script:editMode
    }
    if ($script:deleteRowButton) {
        $script:deleteRowButton.Visible = $script:editMode
    }
    
    # タイトルの更新（ページJSONから読み込む）
    $pageTitle = ""
    $pageConfig = $script:pages[$script:currentPage]
    if ($pageConfig.JsonPath) {
        $pageJsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        }
        else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
        if (Test-Path $pageJsonPath) {
            try {
                $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
                if ($pageJson.Title) {
                    $pageTitle = $pageJson.Title
                }
            }
            catch {
                # エラー時は後続のフォールバック処理に任せる
            }
        }
    }
    # フォールバック: ページJSONにTitleがない場合は、config.jsonまたはデフォルト値を使用
    if (-not $pageTitle) {
        $pageTitle = if ($pageConfig.Title) { $pageConfig.Title } else { if ($script:config.Title) { $script:config.Title } else { "1.V1 移行ツール適用" } }
    }
    $script:titleLabel.Text = $pageTitle
    
    # 移動設定ボタンの表示/非表示とテキストを編集モードに応じて更新
    $currentProcesses = Get-CurrentPageProcesses
    $isPage1 = ($script:currentPage -eq 0)
    $isPage2 = ($script:currentPage -eq 1)
    $isPage3 = ($script:currentPage -eq 2)
    $isPage4 = ($script:currentPage -eq 3)
    for ($i = 0; $i -lt $script:processControls.Count; $i++) {
        $ctrlGroup = $script:processControls[$i]
        if ($ctrlGroup -and $ctrlGroup.FileMoveButton) {
            if ($isPage1) {
                # 1ページ目：常に表示、テキストを編集モードに応じて更新（ONの時は「参照」、OFFの時は「チェック」）
                $ctrlGroup.FileMoveButton.Visible = $true
                if ($script:editMode) {
                    $ctrlGroup.FileMoveButton.Text = "参照"
                }
                else {
                    $ctrlGroup.FileMoveButton.Text = "チェック"  # 設計書通り「チェック」と表示
                }
            }
            elseif ($isPage2) {
                # 2ページ目：常に表示、テキストを編集モードに応じて更新（ONの時は「参照」、OFFの時は「セット」）
                $ctrlGroup.FileMoveButton.Visible = $true
                if ($script:editMode) {
                    $ctrlGroup.FileMoveButton.Text = "参照"
                }
                else {
                    $ctrlGroup.FileMoveButton.Text = "セット"
                }
            }
            elseif ($isPage3) {
                # 3ページ目：常に表示、テキストと色を編集モードに応じて更新（ONの時は「移動設定」水色、OFFの時は「移動」紺色）
                $ctrlGroup.FileMoveButton.Visible = $true
                if ($script:editMode) {
                    $ctrlGroup.FileMoveButton.Text = "移動設定"
                    $ctrlGroup.FileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                    $ctrlGroup.FileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                }
                else {
                    $ctrlGroup.FileMoveButton.Text = "移動"
                    $ctrlGroup.FileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                    $ctrlGroup.FileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                }
            }
            else {
                # その他のページ：編集モードONの時のみ表示
                $ctrlGroup.FileMoveButton.Visible = $script:editMode
            }
        }
        # 3ページ目のCSV名変換ボタンのテキストを編集モードに応じて更新
        if ($isPage3 -and $ctrlGroup -and $ctrlGroup.CsvConvertButton) {
            if ($script:editMode) {
                $ctrlGroup.CsvConvertButton.Text = "参照"
            }
            else {
                $ctrlGroup.CsvConvertButton.Text = "CSV名変換"
            }
        }
        
        # 4ページ目の移動設定ボタンのテキストと色を編集モードに応じて更新
        if ($isPage4 -and $ctrlGroup) {
            # KDL変換CSV格納元の移動設定ボタン（1行目・2行目のみ）
            if ($ctrlGroup.KdlSourceMoveButton) {
                if ($script:editMode) {
                    $ctrlGroup.KdlSourceMoveButton.Text = "移動設定"
                    $ctrlGroup.KdlSourceMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                    $ctrlGroup.KdlSourceMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                }
                else {
                    $ctrlGroup.KdlSourceMoveButton.Text = "移動"
                    $ctrlGroup.KdlSourceMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                    $ctrlGroup.KdlSourceMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                }
            }
            
            # KDL変換CSV格納先の移動設定ボタン（1行目・2行目のみ）
            if ($ctrlGroup.KdlDestMoveButton) {
                if ($script:editMode) {
                    $ctrlGroup.KdlDestMoveButton.Text = "移動設定"
                    $ctrlGroup.KdlDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                    $ctrlGroup.KdlDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                }
                else {
                    $ctrlGroup.KdlDestMoveButton.Text = "移動"
                    $ctrlGroup.KdlDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                    $ctrlGroup.KdlDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                }
            }
            
            # V1抽出CSV格納先の移動設定ボタン（全行）
            if ($ctrlGroup.V1CsvDestMoveButton) {
                if ($script:editMode) {
                    $ctrlGroup.V1CsvDestMoveButton.Text = "移動設定"
                    $ctrlGroup.V1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc（水色）
                    $ctrlGroup.V1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                }
                else {
                    $ctrlGroup.V1CsvDestMoveButton.Text = "移動"
                    $ctrlGroup.V1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 138)  # #1e3a8a（紺色）
                    $ctrlGroup.V1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 40, 100)  # 濃い紺色
                }
            }
        }
        
        # チェックボックスの表示/非表示を編集モードに応じて更新
        if ($ctrlGroup -and $ctrlGroup.CheckBox) {
            $ctrlGroup.CheckBox.Visible = $script:editMode
        }
        
        # プロセス名テキストボックスの編集可否を編集モードに応じて更新
        if ($ctrlGroup -and $ctrlGroup.NameTextBox) {
            $ctrlGroup.NameTextBox.ReadOnly = -not $script:editMode
        }
        
        # 4ページ目のボタンのテキストを編集モードに応じて更新
        if ($isPage4 -and $ctrlGroup) {
            # KDL取込ボタン（1行目・2行目のみ）
            if ($ctrlGroup.KdlImportButton) {
                if ($script:editMode) {
                    $ctrlGroup.KdlImportButton.Text = "参照"
                }
                else {
                    $ctrlGroup.KdlImportButton.Text = "KDL取込"
                }
            }
            
            # 直接取込ボタン（全行）
            if ($ctrlGroup.DirectImportButton) {
                if ($script:editMode) {
                    $ctrlGroup.DirectImportButton.Text = "参照"
                }
                else {
                    $ctrlGroup.DirectImportButton.Text = "直接取込"
                }
            }
            
            # 取込後ボタン（全行）
            if ($ctrlGroup.AfterImportButton) {
                if ($script:editMode) {
                    $ctrlGroup.AfterImportButton.Text = "参照"
                }
                else {
                    if ($i -eq 0) {
                        $ctrlGroup.AfterImportButton.Text = "取込後EA"
                    }
                    else {
                        $ctrlGroup.AfterImportButton.Text = "取込後"
                    }
                }
            }
        }
        
        # ---------------------------------------------------------
        # プロセス無効時・実行済みボタンの表示制御 (Gray-out)
        # ---------------------------------------------------------
        $procEnabled = $true
        $execFlags = @{}
        if ($i -lt $currentProcesses.Count) {
            $pConfig = $currentProcesses[$i]
            if ($pConfig.PSObject.Properties['Enabled']) {
                $procEnabled = $pConfig.Enabled
            }
            # 個別ボタンの実行済みフラグを収集
            foreach ($prop in $pConfig.PSObject.Properties) {
                if ($prop.Name -like "*_Executed") {
                    $execFlags[$prop.Name] = $prop.Value
                }
            }
        }
        
        $grayColor = [System.Drawing.Color]::LightGray
        
        # 1. プロセス全体が無効な場合（編集モードOFF時）
        if (-not $script:editMode -and -not $procEnabled) {
            # テキストボックス
            if ($ctrlGroup.NameTextBox) { $ctrlGroup.NameTextBox.BackColor = $grayColor }
            if ($ctrlGroup.PathTextBox) { $ctrlGroup.PathTextBox.BackColor = $grayColor }
            if ($ctrlGroup.V1CsvSourceTextBox) { $ctrlGroup.V1CsvSourceTextBox.BackColor = $grayColor }
            if ($ctrlGroup.V1CsvDestTextBox) { $ctrlGroup.V1CsvDestTextBox.BackColor = $grayColor }
            if ($ctrlGroup.KdlSourceTextBox) { $ctrlGroup.KdlSourceTextBox.BackColor = $grayColor }
            if ($ctrlGroup.KdlDestTextBox) { $ctrlGroup.KdlDestTextBox.BackColor = $grayColor }
            
            # 実行系ボタン
            if ($ctrlGroup.ExecuteButton) { $ctrlGroup.ExecuteButton.Enabled = $false; $ctrlGroup.ExecuteButton.BackColor = $grayColor }
            if ($ctrlGroup.CsvConvertButton) { $ctrlGroup.CsvConvertButton.Enabled = $false; $ctrlGroup.CsvConvertButton.BackColor = $grayColor }
            if ($ctrlGroup.KdlImportButton) { $ctrlGroup.KdlImportButton.Enabled = $false; $ctrlGroup.KdlImportButton.BackColor = $grayColor }
            if ($ctrlGroup.KdlKdbButton) { $ctrlGroup.KdlKdbButton.Enabled = $false; $ctrlGroup.KdlKdbButton.BackColor = $grayColor }
            if ($ctrlGroup.KdlEbButton) { $ctrlGroup.KdlEbButton.Enabled = $false; $ctrlGroup.KdlEbButton.BackColor = $grayColor }
            if ($ctrlGroup.DirectImportButton) { $ctrlGroup.DirectImportButton.Enabled = $false; $ctrlGroup.DirectImportButton.BackColor = $grayColor }
            if ($ctrlGroup.AfterImportButton) { $ctrlGroup.AfterImportButton.Enabled = $false; $ctrlGroup.AfterImportButton.BackColor = $grayColor }
            if ($ctrlGroup.MaintButton) { $ctrlGroup.MaintButton.Enabled = $false; $ctrlGroup.MaintButton.BackColor = $grayColor }
            if ($ctrlGroup.MaintButton1) { $ctrlGroup.MaintButton1.Enabled = $false; $ctrlGroup.MaintButton1.BackColor = $grayColor }
            if ($ctrlGroup.MaintButton2) { $ctrlGroup.MaintButton2.Enabled = $false; $ctrlGroup.MaintButton2.BackColor = $grayColor }
        }
        # 2. 個別ボタンが実行済みの場合（プロセスが有効であっても非活性化）
        elseif (-not $script:editMode) {
            if ($ctrlGroup.ExecuteButton -and $execFlags["ExecuteButton_Executed"]) { 
                $ctrlGroup.ExecuteButton.Enabled = $false; $ctrlGroup.ExecuteButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.CsvConvertButton -and $execFlags["CsvConvertButton_Executed"]) { 
                $ctrlGroup.CsvConvertButton.Enabled = $false; $ctrlGroup.CsvConvertButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.KdlImportButton -and $execFlags["KdlImportButton_Executed"]) { 
                $ctrlGroup.KdlImportButton.Enabled = $false; $ctrlGroup.KdlImportButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.KdlKdbButton -and $execFlags["KdlKdbButton_Executed"]) { 
                $ctrlGroup.KdlKdbButton.Enabled = $false; $ctrlGroup.KdlKdbButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.KdlEbButton -and $execFlags["KdlEbButton_Executed"]) { 
                $ctrlGroup.KdlEbButton.Enabled = $false; $ctrlGroup.KdlEbButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.DirectImportButton -and $execFlags["DirectImportButton_Executed"]) { 
                $ctrlGroup.DirectImportButton.Enabled = $false; $ctrlGroup.DirectImportButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.AfterImportButton -and $execFlags["AfterImportButton_Executed"]) { 
                $ctrlGroup.AfterImportButton.Enabled = $false; $ctrlGroup.AfterImportButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.MaintButton -and $execFlags["MaintButton_Executed"]) { 
                $ctrlGroup.MaintButton.Enabled = $false; $ctrlGroup.MaintButton.BackColor = $grayColor 
            }
            if ($ctrlGroup.MaintButton1 -and $execFlags["MaintButton1_Executed"]) { 
                $ctrlGroup.MaintButton1.Enabled = $false; $ctrlGroup.MaintButton1.BackColor = $grayColor 
            }
            if ($ctrlGroup.MaintButton2 -and $execFlags["MaintButton2_Executed"]) { 
                $ctrlGroup.MaintButton2.Enabled = $false; $ctrlGroup.MaintButton2.BackColor = $grayColor 
            }
        }
    }
    
    # ページパスの読み込み
    Update-PagePaths
    
    # ログ集約ボタンのテキストを編集モードに応じて更新
    if ($script:logAggregationButton) {
        if ($script:editMode) {
            $script:logAggregationButton.Text = "参照"
        }
        else {
            $script:logAggregationButton.Text = "集約"
        }
    }
    
    # ログ格納ボタンのテキストを編集モードに応じて更新
    if ($script:logStorageButton) {
        if ($script:editMode) {
            $script:logStorageButton.Text = "参照"
        }
        else {
            $script:logStorageButton.Text = "ログ格納"
        }
    }
    
    # ページに応じてレイアウトを調整
    if ($useDrawioLayout) {
        # 1ページ目・2ページ目：ファイル移動セクションを非表示
        if ($script:headerPanel) {
            if ($isPage2) {
                # 2ページ目：薄いオレンジ色
                $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)
            }
            else {
                # 1ページ目：薄い青色
                $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
            }
        }

        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # ログ格納セクションの位置を調整（370px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 370)
        }
        

        
        # ログ出力エリアの位置を調整（430px y座標、740px幅、130px高さ）
        if ($script:logTextBox) {
            $script:logTextBox.Location = New-Object System.Drawing.Point(10, 430)
            $script:logTextBox.Size = New-Object System.Drawing.Size(740, 130)
        }
        
        # フォームの高さを調整（600px）
        if ($script:form) {
            $script:form.Size = New-Object System.Drawing.Size(900, 600)
        }
        
        # プロセスパネルの高さを調整（320px）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 320)
        }
    }
    elseif ($isPage3) {
        # 3ページ目：JAVA移行ツール実行のレイアウト
        # ヘッダーの背景色を緑色に変更
        if ($script:headerPanel) {
            $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(147, 196, 125)  # #93C47D
        }
        
        # ファイル移動セクションを非表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # プロセスパネルの高さを調整（300px：プロセス3つの下に余裕を持たせる）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 300)
        }
        
        # ログ格納セクションの位置を調整（プロセスパネルの下：50 + 300 = 350px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 350)
        }
        

        
        # ログ出力エリアの位置を調整（ログ格納セクションの下：350 + 60 = 410px y座標、740px幅、130px高さ）
        if ($script:logTextBox) {
            $script:logTextBox.Location = New-Object System.Drawing.Point(10, 410)
            $script:logTextBox.Size = New-Object System.Drawing.Size(740, 130)
        }
        
        # フォームの高さを調整（ログ出力エリアの下：410 + 130 = 540px、余裕を持たせて600px）
        if ($script:form) {
            $script:form.Size = New-Object System.Drawing.Size(900, 600)
        }
    }
    elseif ($isPage4) {
        # 4ページ目：SQLLOADER実行のレイアウト
        # ヘッダーの背景色を青色に変更
        if ($script:headerPanel) {
            $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(27, 161, 226)  # #1ba1e2
        }
        
        # ファイル移動セクションを非表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # プロセスパネルの高さを調整（320pxに短縮：1,2ページと同じ）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 320)
        }
        
        # ログ格納セクションを表示
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Visible = $true
        }
        
        # ログ格納セクションの位置を調整（プロセスパネルの下：50 + 320 = 370px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 370)
        }
        

        
        # ログ出力エリアの位置を調整（ログ格納セクションの下：370 + 60 = 430px y座標、740px幅、130px高さ）
        if ($script:logTextBox) {
            $script:logTextBox.Location = New-Object System.Drawing.Point(10, 430)
            $script:logTextBox.Size = New-Object System.Drawing.Size(740, 130)
        }
        
        # フォームの高さを調整（600px：他ページと同じ）
        if ($script:form) {
            $script:form.Size = New-Object System.Drawing.Size(900, 600)
        }
    }
    else {
        # 5ページ目以降：従来のレイアウト
        # ヘッダーの背景色を水色に戻す
        if ($script:headerPanel) {
            $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
        }
        
        # ファイル移動セクションを表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $true
        }
        
        # ログ格納セクションを表示
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Visible = $true
        }
        
        # ログ格納セクションの位置を元に戻す（490px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 490)
        }

        
        # ログ出力エリアの位置を元に戻す（605px y座標、880px幅、220px高さ）
        if ($script:logTextBox) {
            $script:logTextBox.Location = New-Object System.Drawing.Point(10, 605)
            $script:logTextBox.Size = New-Object System.Drawing.Size(880, 220)
        }
        
        # フォームの高さを元に戻す（1000px）
        if ($script:form) {
            $script:form.Size = New-Object System.Drawing.Size(900, 1000)
        }
        
        # プロセスパネルの高さを元に戻す（320px）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 320)
        }
    }
}
