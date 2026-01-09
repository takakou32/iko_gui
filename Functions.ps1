# PowerShellスクリプト - 関数定義
# エンコーディング: UTF-8 BOM付

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
    param([string]$Message, [string]$Level = "INFO", [int]$ProcessIndex = -1, [string]$LogDir = $null)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # プロセス固有のログファイル
    if ($ProcessIndex -ge 0) {
        # LogDirが指定されていない場合、ProcessIndexから取得
        if (-not $LogDir) {
            $currentProcesses = Get-CurrentPageProcesses
            if ($currentProcesses -and $ProcessIndex -lt $currentProcesses.Count) {
                $processConfig = $currentProcesses[$ProcessIndex]
                if ($processConfig.LogOutputDir) {
                    $LogDir = if ([System.IO.Path]::IsPathRooted($processConfig.LogOutputDir)) {
                        $processConfig.LogOutputDir
                    } else {
                        Join-Path $PSScriptRoot $processConfig.LogOutputDir
                    }
                } else {
                    $LogDir = $script:logDir
                }
            } else {
                $LogDir = $script:logDir
            }
        }
        
        # ログディレクトリが存在しない場合は作成
        if (-not (Test-Path $LogDir)) {
            New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
        }
        
        $processLogFile = Join-Path $LogDir "process_${script:currentPage}_${ProcessIndex}.log"
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
    
    # パスの正規化処理
    if ([string]::IsNullOrWhiteSpace($BatchPath)) {
        Write-Log "バッチファイルパスが空です" "ERROR" $ProcessIndex
        return $false
    }
    
    # 先頭・末尾の空白を削除
    $BatchPath = $BatchPath.Trim()
    
    # パスを正規化（相対パスの解決、区切り文字の統一など）
    try {
        # 相対パスの場合は$PSScriptRootを基準に解決
        if (-not [System.IO.Path]::IsPathRooted($BatchPath)) {
            $BatchPath = Join-Path $PSScriptRoot $BatchPath
        }
        # パスを正規化（..や.を解決、区切り文字を統一）
        $BatchPath = [System.IO.Path]::GetFullPath($BatchPath)
    } catch {
        Write-Log "バッチファイルパスの正規化に失敗しました: $BatchPath - $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
    
    if (-not (Test-Path $BatchPath)) {
        Write-Log "バッチファイルが見つかりません: $BatchPath" "ERROR" $ProcessIndex
        return $false
    }
    
    # ログ出力ディレクトリの決定
    $currentProcesses = Get-CurrentPageProcesses
    $processLogDir = $script:logDir
    if ($currentProcesses -and $ProcessIndex -lt $currentProcesses.Count) {
        $processConfig = $currentProcesses[$ProcessIndex]
        if ($processConfig.LogOutputDir) {
            $processLogDir = if ([System.IO.Path]::IsPathRooted($processConfig.LogOutputDir)) {
                $processConfig.LogOutputDir
            } else {
                Join-Path $PSScriptRoot $processConfig.LogOutputDir
            }
            if (-not (Test-Path $processLogDir)) {
                New-Item -ItemType Directory -Path $processLogDir -Force | Out-Null
            }
        }
    }
    
    Write-Log "バッチファイルを実行中: $DisplayName ($BatchPath)" "INFO" $ProcessIndex
    
    try {
        $stdoutFile = Join-Path $processLogDir "process_${script:currentPage}_${ProcessIndex}_stdout.log"
        $stderrFile = Join-Path $processLogDir "process_${script:currentPage}_${ProcessIndex}_stderr.log"
        
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

# バッチファイルパス保存関数
function Save-BatchFilePath {
    param([int]$ProcessIndex, [string]$BatchFilePath, [int]$BatchIndex = 0)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    } else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $jsonPath)) {
        Write-Log "JSONファイルが見つかりません: $jsonPath" "ERROR" $ProcessIndex
        return $false
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not $jsonContent.Processes -or $ProcessIndex -ge $jsonContent.Processes.Count) {
            Write-Log "プロセスインデックスが範囲外です" "ERROR" $ProcessIndex
            return $false
        }
        
        $process = $jsonContent.Processes[$ProcessIndex]
        if (-not $process.BatchFiles) {
            $process.BatchFiles = @()
        }
        
        if ($BatchIndex -ge $process.BatchFiles.Count) {
            # 新しいバッチファイルエントリを追加
            $process.BatchFiles += @{
                Name = "バッチファイル"
                Path = $BatchFilePath
            }
        } else {
            # 既存のバッチファイルエントリを更新
            $process.BatchFiles[$BatchIndex].Path = $BatchFilePath
        }
        
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
            $targetPath = [System.IO.Path]::GetFullPath($BatchFilePath).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                if ([string]::IsNullOrEmpty($relative)) {
                    $relative = Split-Path $targetPath -Leaf
                }
                $relative
            } else {
                $BatchFilePath
            }
        } catch {
            $BatchFilePath
        }
        
        $process.BatchFiles[$BatchIndex].Path = $relativePath
        
        # JSONファイルに保存
        $jsonContent | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
        Write-Log "バッチファイルパスを保存しました: $relativePath" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ログ出力フォルダパス保存関数
function Save-ProcessLogOutputDir {
    param([int]$ProcessIndex, [string]$LogOutputDir)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    } else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $jsonPath)) {
        Write-Log "JSONファイルが見つかりません: $jsonPath" "ERROR" $ProcessIndex
        return $false
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not $jsonContent.Processes -or $ProcessIndex -ge $jsonContent.Processes.Count) {
            Write-Log "プロセスインデックスが範囲外です" "ERROR" $ProcessIndex
            return $false
        }
        
        $process = $jsonContent.Processes[$ProcessIndex]
        
        # LogOutputDirプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name "LogOutputDir" -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name "LogOutputDir" -Value $LogOutputDir
        } else {
            $process.LogOutputDir = $LogOutputDir
        }
        
        # JSONファイルに保存
        $jsonContent | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
        Write-Log "ログ出力フォルダパスを保存しました: $LogOutputDir" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ページパス読み込み関数
function Load-PagePaths {
    if ($script:currentPage -ge $script:pages.Count) {
        return
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # ページJSONファイルから設定を読み込む
    $pageJsonPath = $null
    if ($pageConfig.JsonPath) {
        $pageJsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        } else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
    }
    
    $sourcePath = ""
    $destPath = ""
    $logStoragePath = ""
    
    # ページJSONファイルが存在する場合はそこから読み込む
    if ($pageJsonPath -and (Test-Path $pageJsonPath)) {
        try {
            $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
            $sourcePath = if ($pageJson.SourcePath) { $pageJson.SourcePath } else { "" }
            $destPath = if ($pageJson.DestinationPath) { $pageJson.DestinationPath } else { "" }
            $logStoragePath = if ($pageJson.LogStoragePath) { $pageJson.LogStoragePath } else { "" }
        } catch {
            Write-Log "ページJSONファイルの読み込みに失敗しました: $pageJsonPath - $($_.Exception.Message)" "ERROR"
        }
    }
    
    # 移行データファイル移動元
    if ($sourcePath -and $sourcePath -ne "パス" -and $sourcePath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($sourcePath)) {
                $sourcePath = Join-Path $PSScriptRoot $sourcePath
            }
            $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
            $script:sourcePathTextBox.Text = $sourcePath
        } catch {
            $script:sourcePathTextBox.Text = "パス"
        }
    } else {
        $script:sourcePathTextBox.Text = "パス"
    }
    
    # 移行データファイル移動先
    if ($destPath -and $destPath -ne "パス" -and $destPath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($destPath)) {
                $destPath = Join-Path $PSScriptRoot $destPath
            }
            $destPath = [System.IO.Path]::GetFullPath($destPath)
            $script:destPathTextBox.Text = $destPath
        } catch {
            $script:destPathTextBox.Text = "パス"
        }
    } else {
        $script:destPathTextBox.Text = "パス"
    }
    
    # ログ格納先
    if ($logStoragePath -and $logStoragePath -ne "パス" -and $logStoragePath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($logStoragePath)) {
                $logStoragePath = Join-Path $PSScriptRoot $logStoragePath
            }
            $logStoragePath = [System.IO.Path]::GetFullPath($logStoragePath)
            $script:logStoragePathTextBox.Text = $logStoragePath
        } catch {
            $script:logStoragePathTextBox.Text = "パス"
        }
    } else {
        $script:logStoragePathTextBox.Text = "パス"
    }
}

# ページパス保存関数
function Save-PagePaths {
    param(
        [string]$SourcePath = $null,
        [string]$DestinationPath = $null,
        [string]$LogStoragePath = $null
    )
    
    if ($script:currentPage -ge $script:pages.Count) {
        Write-Log "ページインデックスが範囲外です" "ERROR"
        return $false
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # ページJSONファイルのパスを取得
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN"
        return $false
    }
    
    $pageJsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    } else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $pageJsonPath)) {
        Write-Log "ページJSONファイルが見つかりません: $pageJsonPath" "ERROR"
        return $false
    }
    
    try {
        # ページJSONファイルを読み込む
        $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # 相対パスに変換（可能な場合）
        if ($SourcePath) {
            $relativeSourcePath = try {
                $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                $targetPath = [System.IO.Path]::GetFullPath($SourcePath).TrimEnd('\', '/')
                
                if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                    if ([string]::IsNullOrEmpty($relative)) {
                        $relative = Split-Path $targetPath -Leaf
                    }
                    $relative
                } else {
                    $SourcePath
                }
            } catch {
                $SourcePath
            }
            $pageJson.SourcePath = $relativeSourcePath
        }
        
        if ($DestinationPath) {
            $relativeDestPath = try {
                $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                $targetPath = [System.IO.Path]::GetFullPath($DestinationPath).TrimEnd('\', '/')
                
                if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                    if ([string]::IsNullOrEmpty($relative)) {
                        $relative = Split-Path $targetPath -Leaf
                    }
                    $relative
                } else {
                    $DestinationPath
                }
            } catch {
                $DestinationPath
            }
            $pageJson.DestinationPath = $relativeDestPath
        }
        
        if ($LogStoragePath) {
            $relativeLogPath = try {
                $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                $targetPath = [System.IO.Path]::GetFullPath($LogStoragePath).TrimEnd('\', '/')
                
                if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                    if ([string]::IsNullOrEmpty($relative)) {
                        $relative = Split-Path $targetPath -Leaf
                    }
                    $relative
                } else {
                    $LogStoragePath
                }
            } catch {
                $LogStoragePath
            }
            $pageJson.LogStoragePath = $relativeLogPath
        }
        
        # ページJSONファイルに保存
        $pageJson | ConvertTo-Json -Depth 10 | Set-Content $pageJsonPath -Encoding UTF8
        
        Write-Log "ページパスを保存しました: $pageJsonPath" "INFO"
        return $true
    } catch {
        Write-Log "ページJSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# プロセス実行関数
function Start-ProcessFlow {
    param([int]$ProcessIndex)
    
    # 編集モード中はファイル選択ダイアログを表示
    if ($script:editMode) {
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
        $fileDialog.Title = "バッチファイルを選択してください"
        
        # 現在のバッチファイルパスを初期値として設定
        $currentProcesses = Get-CurrentPageProcesses
        $processConfig = $currentProcesses[$ProcessIndex]
        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
            $currentBatch = $processConfig.BatchFiles[0]
            $initialPath = if ([System.IO.Path]::IsPathRooted($currentBatch.Path)) {
                $currentBatch.Path
            } else {
                Join-Path $PSScriptRoot $currentBatch.Path
            }
            if (Test-Path $initialPath) {
                $fileDialog.InitialDirectory = Split-Path $initialPath
                $fileDialog.FileName = Split-Path $initialPath -Leaf
            }
        }
        
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $fileDialog.FileName
            Save-BatchFilePath -ProcessIndex $ProcessIndex -BatchFilePath $selectedFile -BatchIndex 0
            Write-Log "バッチファイルを設定しました: $selectedFile" "INFO" $ProcessIndex
            [System.Windows.Forms.MessageBox]::Show("バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # コントロールを更新して新しい設定を反映
            Update-ProcessControls
        }
        $fileDialog.Dispose()
        return
    }
    
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
    
    
    if ($allSuccess) {
        Write-Log "プロセスが正常に完了しました: $($processConfig.Name)" "INFO" $ProcessIndex
    } else {
        Write-Log "プロセスでエラーが発生しました: $($processConfig.Name)" "ERROR" $ProcessIndex
    }
    
    $executeButton.Enabled = $true
}

# ファイル移動設定ダイアログ表示関数
function Show-FileMoveSettingsDialog {
    param([int]$ProcessIndex, [string]$ProcessName)
    
    # 現在のプロセス設定を取得
    $currentProcesses = Get-CurrentPageProcesses
    if (-not $currentProcesses -or $ProcessIndex -ge $currentProcesses.Count) {
        [System.Windows.Forms.MessageBox]::Show("プロセス情報を取得できませんでした。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $processConfig = $currentProcesses[$ProcessIndex]
    [System.Windows.Forms.MessageBox]::Show("デバッグProcessName: $ProcessName")
    # movefiles フォルダのファイルを読み込み
    $initialText = ""
    $fileNameRaw = if ($processConfig.Name) { $processConfig.Name } elseif ($ProcessName) { $ProcessName } else { "" }   
    $fileNameRaw = $fileNameRaw.Trim()
    if (-not $fileNameRaw) {
        $fileNameRaw = "process_${script:currentPage + 1}_${ProcessIndex + 1}"
    }
    $safeFileName = [regex]::Replace($fileNameRaw, '[<>:"/\\|?*\r\n\t]', '_')
    $moveFilesDir = Join-Path $PSScriptRoot "movefiles"
    if (Test-Path $moveFilesDir) {
        $candidatePath = Join-Path $moveFilesDir ($safeFileName + ".txt")
        if (Test-Path $candidatePath) {
            try {
                $initialText = Get-Content -Path $candidatePath -Encoding UTF8 -Raw
            } catch {
                # ファイル読み込みエラーは無視
            }
        }
    }
    
    # ダイアログフォームを作成
    $dialogForm = New-Object System.Windows.Forms.Form
    $dialogForm.Text = "ファイル移動設定 - $($processConfig.Name)"
    $dialogForm.Size = New-Object System.Drawing.Size(600, 400)
    $dialogForm.StartPosition = "CenterParent"
    $dialogForm.FormBorderStyle = "FixedDialog"
    $dialogForm.MaximizeBox = $false
    $dialogForm.MinimizeBox = $false
    $dialogForm.ShowInTaskbar = $false
    $dialogForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # 説明ラベル
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(560, 40)
    $label.Text = "移動元パス|移動先パス の形式で1行に1つずつ入力してください。`n例: csv_source|csv_destination"
    $label.Font = New-Object System.Drawing.Font("メイリオ", 9)
    $dialogForm.Controls.Add($label)
    
    # テキスト入力エリア
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 55)
    $textBox.Size = New-Object System.Drawing.Size(560, 250)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
    $textBox.Text = $initialText
    $textBox.AcceptsReturn = $true  # エンターキーで改行できるようにする
    $dialogForm.Controls.Add($textBox)
    
    # 保存ボタン
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(400, 320)
    $saveButton.Size = New-Object System.Drawing.Size(80, 30)
    $saveButton.Text = "保存"
    $saveButton.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 255)
    $saveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $saveButton.FlatAppearance.BorderSize = 1
    $saveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
    $saveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # AcceptButtonを設定しない（エンターキーで改行できるようにするため）
    $dialogForm.Controls.Add($saveButton)
    
    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(490, 320)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "キャンセル"
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialogForm.CancelButton = $cancelButton
    $dialogForm.Controls.Add($cancelButton)
    
    # ダイアログを表示
    $result = $dialogForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # テキストをファイルに保存（movefiles フォルダ）
        # ボタン左隣のテキストボックス内の文字列（ProcessName）をファイル名として使用
        $fileNameRaw = if ($ProcessName) { $ProcessName } else { "" }
        $fileNameRaw = $fileNameRaw.Trim()
        if (-not $fileNameRaw) {
            # 空の場合はページ・インデックスで代替
            $fileNameRaw = "process_${script:currentPage + 1}_${ProcessIndex + 1}"
        }
        # ファイル名として使えない文字を置換（改行やタブも除去）
        $safeFileName = [regex]::Replace($fileNameRaw, '[<>:"/\\|?*\r\n\t]', '_')
        $moveFilesDir = Join-Path $PSScriptRoot "movefiles"
        if (-not (Test-Path $moveFilesDir)) {
            New-Item -ItemType Directory -Path $moveFilesDir -Force | Out-Null
        }
        $moveFilePath = Join-Path $moveFilesDir ($safeFileName + ".txt")
        Set-Content -Path $moveFilePath -Value $textBox.Text -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("ファイル移動設定を保存しました。", "保存完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    $dialogForm.Dispose()
}

# ログ確認関数
function Show-ProcessLog {
    param([int]$ProcessIndex)
    
    # 編集モード中はフォルダ選択ダイアログを表示
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "ログ出力フォルダを選択してください"
        $folderDialog.ShowNewFolderButton = $true
        
        # 現在のログフォルダを初期値として設定
        $currentProcesses = Get-CurrentPageProcesses
        $processConfig = $currentProcesses[$ProcessIndex]
        if ($processConfig.LogOutputDir) {
            $initialPath = if ([System.IO.Path]::IsPathRooted($processConfig.LogOutputDir)) {
                $processConfig.LogOutputDir
            } else {
                Join-Path $PSScriptRoot $processConfig.LogOutputDir
            }
            if (Test-Path $initialPath) {
                $folderDialog.SelectedPath = $initialPath
            }
        } else {
            if (Test-Path $script:logDir) {
                $folderDialog.SelectedPath = $script:logDir
            }
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            # 相対パスに変換（可能な場合）
            $relativePath = try {
                $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                $targetPath = [System.IO.Path]::GetFullPath($selectedPath).TrimEnd('\', '/')
                
                if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                    if ([string]::IsNullOrEmpty($relative)) {
                        $relative = Split-Path $targetPath -Leaf
                    }
                    $relative
                } else {
                    $selectedPath
                }
            } catch {
                $selectedPath
            }
            
            # JSONファイルを更新
            if (Save-ProcessLogOutputDir -ProcessIndex $ProcessIndex -LogOutputDir $relativePath) {
                Write-Log "ログ出力フォルダを設定しました: $relativePath" "INFO" $ProcessIndex
                [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダを設定しました。`n$relativePath", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダの保存に失敗しました。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        $folderDialog.Dispose()
        return
    }
    
    # 通常モードではJSONファイルで設定されているログ出力フォルダをエクスプローラで開く
    # JSONファイルから最新の情報を読み込む
    $currentProcesses = Get-CurrentPageProcesses
    if (-not $currentProcesses -or $ProcessIndex -ge $currentProcesses.Count) {
        [System.Windows.Forms.MessageBox]::Show("プロセス情報を取得できませんでした。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Write-Log "プロセス情報を取得できませんでした: ProcessIndex=$ProcessIndex" "ERROR" $ProcessIndex
        return
    }
    
    $processConfig = $currentProcesses[$ProcessIndex]
    
    # JSONファイルで設定されているLogOutputDirを取得
    $processLogDir = $logDir  # デフォルト値
    if ($processConfig.LogOutputDir) {
        # LogOutputDirが設定されている場合はそれを使用
        $processLogDir = if ([System.IO.Path]::IsPathRooted($processConfig.LogOutputDir)) {
            $processConfig.LogOutputDir
        } else {
            Join-Path $PSScriptRoot $processConfig.LogOutputDir
        }
    }
    
    # ログ出力フォルダをエクスプローラで開く
    if (Test-Path $processLogDir) {
        # エクスプローラでフォルダを開く
        try {
            Start-Process explorer.exe -ArgumentList $processLogDir
            Write-Log "ログ出力フォルダを開きました: $processLogDir" "INFO" $ProcessIndex
        } catch {
            [System.Windows.Forms.MessageBox]::Show("エクスプローラを起動できませんでした。`n$processLogDir`n`n$($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "エクスプローラを起動できませんでした: $processLogDir - $($_.Exception.Message)" "ERROR" $ProcessIndex
        }
    } else {
        # フォルダが存在しない場合は作成してから開く
        try {
            New-Item -ItemType Directory -Path $processLogDir -Force | Out-Null
            Start-Process explorer.exe -ArgumentList $processLogDir
            Write-Log "ログ出力フォルダを作成して開きました: $processLogDir" "INFO" $ProcessIndex
        } catch {
            [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダを開けませんでした。`n$processLogDir`n`n$($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "ログ出力フォルダを開けませんでした: $processLogDir - $($_.Exception.Message)" "ERROR" $ProcessIndex
        }
    }
}

# プロセスコントロールの更新
function Update-ProcessControls {
    # ページ遷移時にJSONファイルを読み込む
    $currentProcesses = Get-CurrentPageProcesses
    $totalPages = $script:pages.Count
    
    Write-Log "ページ $($script:currentPage + 1) のプロセスを読み込みました (プロセス数: $($currentProcesses.Count))" "INFO"
    
    # 既存のコントロールをすべてクリア（processPanel内のすべてのコントロールを削除）
    # まず、processControls配列に保存されているコントロールを削除
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup) {
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
        } catch {
            # エラーは無視（既に削除されている可能性がある）
        }
    }
    
    # 念のため、processPanel.Controlsをクリア（すべてのコントロールを削除）
    # これにより、前のページのコントロールが確実に削除される
    $script:processPanel.Controls.Clear()
    
    # ページ番号を判定（drawioのレイアウトを適用）
    $isPage1 = ($script:currentPage -eq 0)
    $isPage2 = ($script:currentPage -eq 1)
    $isPage3 = ($script:currentPage -eq 2)
    $isPage4 = ($script:currentPage -eq 3)
    $useDrawioLayout = ($isPage1 -or $isPage2)
    
    # 新しいコントロールを作成
    for ($i = 0; $i -lt $script:processesPerPage; $i++) {
        if ($i -lt $currentProcesses.Count) {
            $processConfig = $currentProcesses[$i]
            $row = [Math]::Floor($i / 2)
            $col = $i % 2
            
            if ($useDrawioLayout) {
                # 1ページ目・2ページ目：drawioのレイアウトに合わせる
                # drawioの座標: タスク名(60, 100+), セット/チェック(200, 100+), 実行(280, 100+), ログ確認(350, 100+)
                # プロセスパネルのy座標は50なので、実際のy座標は50から（100-50=50）
                $x = if ($col -eq 0) { 60 } else { 440 }
                $y = 50 + $row * 40
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = $true
                $nameTextBox.BackColor = [System.Drawing.Color]::White
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
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
                    } else {
                        $fileMoveButton.Text = "チェック"  # 設計書通り「チェック」と表示
                    }
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc（設計書通りの赤色）
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450（設計書通りの赤色ボーダー）
                    $fileMoveButton.Visible = $true  # 常に表示
                    $processIdx = $i
                    $fileMoveButton.Add_Click({
                        Start-ProcessFlow -ProcessIndex $processIdx
                    })
                } else {
                    # 2ページ目：セットボタン
                    # 編集モードOFF時は「セット」と表示し、実行ボタンと同じ機能（プロセス実行）
                    # 編集モードON時は「参照」と表示し、実行ボタンの編集モードON時と同じ機能（ファイル選択ウィザードを開き、パスをJSONに保存）
                    if ($script:editMode) {
                        $fileMoveButton.Text = "参照"
                    } else {
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
                            $fileDialog.Title = "バッチファイルを選択してください"
                            
                            # 現在のバッチファイルパスを初期値として設定
                            $currentProcesses = Get-CurrentPageProcesses
                            if ($currentProcesses -and $clickedProcessIdx -lt $currentProcesses.Count) {
                                $processConfig = $currentProcesses[$clickedProcessIdx]
                                if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
                                    $currentBatch = $processConfig.BatchFiles[0]
                                    $initialPath = if ([System.IO.Path]::IsPathRooted($currentBatch.Path)) {
                                        $currentBatch.Path
                                    } else {
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
                                Save-BatchFilePath -ProcessIndex $clickedProcessIdx -BatchFilePath $selectedFile -BatchIndex 0
                                Write-Log "バッチファイルを設定しました: $selectedFile" "INFO" $clickedProcessIdx
                                [System.Windows.Forms.MessageBox]::Show("バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                
                                # コントロールを更新して新しい設定を反映
                                Update-ProcessControls
                            }
                            $fileDialog.Dispose()
                        } else {
                            # 編集モードOFF時は実行ボタンと同じ機能（プロセス実行）
                            Start-ProcessFlow -ProcessIndex $clickedProcessIdx
                        }
                    })
                    $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc
                    $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450
                    $fileMoveButton.Visible = $true  # 編集モードOFF時も表示
                }
                $fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $fileMoveButton.FlatAppearance.BorderSize = 1
                $fileMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $script:processPanel.Controls.Add($fileMoveButton)
                
                # 実行ボタン（オレンジ）
                $executeButton = New-Object System.Windows.Forms.Button
                $executeX = $x + 220
                $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
                $executeButton.Size = New-Object System.Drawing.Size(60, 30)
                if ($script:editMode) {
                    $executeButton.Text = "参照"
                } else {
                    $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                }
                # 実行ボタンの見た目（設計書通り：fillColor=#ffcc99, strokeColor=#d6b656）
                $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $executeButton.FlatAppearance.BorderSize = 1
                $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $processIdx = $i
                $executeButton.Add_Click({
                    Start-ProcessFlow -ProcessIndex $processIdx
                })
                $script:processPanel.Controls.Add($executeButton)
                
                # ログ確認ボタン（緑）
                $logButton = New-Object System.Windows.Forms.Button
                $logX = $x + 290
                $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                $logButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($script:editMode) {
                    $logButton.Text = "参照"
                } else {
                    $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                }
                $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                $logButton.FlatAppearance.BorderSize = 1
                $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $processIdx = $i
                $logButton.Add_Click({
                    Show-ProcessLog -ProcessIndex $processIdx
                })
                $script:processPanel.Controls.Add($logButton)
                
                # 1ページ目・2ページ目用のコントロール情報を保存
                $script:processControls += @{
                    NameTextBox = $nameTextBox
                    FileMoveButton = $fileMoveButton
                    ExecuteButton = $executeButton
                    LogButton = $logButton
                }
            } elseif ($isPage3) {
                # 3ページ目：JAVA移行ツール実行のレイアウト
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
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "V1抽出CSV格納元フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $v1CsvSourceTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
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
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = $true
                $nameTextBox.BackColor = [System.Drawing.Color]::White
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
                $script:processPanel.Controls.Add($nameTextBox)
                
                # パス入力テキストボックス（V1抽出CSV格納先）
                $pathTextBox = New-Object System.Windows.Forms.TextBox
                $pathX = 210
                $pathTextBox.Location = New-Object System.Drawing.Point($pathX, $y)
                $pathTextBox.Size = New-Object System.Drawing.Size(220, 30)
                $pathTextBox.Text = "パス"
                $pathTextBox.ReadOnly = $true
                $pathTextBox.BackColor = [System.Drawing.Color]::White
                $pathTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $pathTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $pathTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                $pathTextBox.Add_Click({
                    if ($script:editMode) {
                        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                        $folderDialog.Description = "V1抽出CSV格納先フォルダを選択してください"
                        $folderDialog.ShowNewFolderButton = $true
                        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $pathTextBox.Text = $folderDialog.SelectedPath
                        }
                        $folderDialog.Dispose()
                    }
                })
                $script:processPanel.Controls.Add($pathTextBox)
                
                # 移動設定ボタン（水色）
                $fileMoveButton = New-Object System.Windows.Forms.Button
                $fileMoveX = 440
                $fileMoveButton.Location = New-Object System.Drawing.Point($fileMoveX, $y)
                $fileMoveButton.Size = New-Object System.Drawing.Size(70, 30)
                $fileMoveButton.Text = "移動設定"
                $fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc
                $fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
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
                
                # CSV名変換ボタン（赤色）- 機能は未実装のため、見た目のみ
                $csvConvertButton = New-Object System.Windows.Forms.Button
                $csvConvertX = 520
                $csvConvertButton.Location = New-Object System.Drawing.Point($csvConvertX, $y)
                $csvConvertButton.Size = New-Object System.Drawing.Size(80, 30)
                $csvConvertButton.Text = "CSV名変換"
                $csvConvertButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc
                $csvConvertButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $csvConvertButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450
                $csvConvertButton.FlatAppearance.BorderSize = 1
                $csvConvertButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $csvConvertButton.Enabled = $false  # 機能未実装
                $script:processPanel.Controls.Add($csvConvertButton)
                
                # 実行ボタン（オレンジ）
                $executeButton = New-Object System.Windows.Forms.Button
                $executeX = 610
                $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
                $executeButton.Size = New-Object System.Drawing.Size(60, 30)
                if ($script:editMode) {
                    $executeButton.Text = "参照"
                } else {
                    $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                }
                $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                $executeButton.FlatAppearance.BorderSize = 1
                $executeButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $processIdx = $i
                $executeButton.Add_Click({
                    Start-ProcessFlow -ProcessIndex $processIdx
                })
                $script:processPanel.Controls.Add($executeButton)
                
                # ログ確認ボタン（緑）
                $logButton = New-Object System.Windows.Forms.Button
                $logX = 680
                $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                $logButton.Size = New-Object System.Drawing.Size(70, 30)
                if ($script:editMode) {
                    $logButton.Text = "参照"
                } else {
                    $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                }
                $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                $logButton.FlatAppearance.BorderSize = 1
                $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $processIdx = $i
                $logButton.Add_Click({
                    Show-ProcessLog -ProcessIndex $processIdx
                })
                $script:processPanel.Controls.Add($logButton)
                
                # 3ページ目用のコントロール情報を保存
                $script:processControls += @{
                    NameTextBox = $nameTextBox
                    PathTextBox = $pathTextBox
                    FileMoveButton = $fileMoveButton
                    CsvConvertButton = $csvConvertButton
                    ExecuteButton = $executeButton
                    LogButton = $logButton
                }
            } elseif ($isPage4) {
                # 4ページ目：SQLLOADER実行のレイアウト
                # drawioの座標を参考に、プロセスパネルのy座標50を考慮
                # 1行目: タスク名(30, 190)、KDL変換CSV格納元(170, 190)、KDL変換CSV格納先(510, 190)、V1抽出CSV格納先(510, 245)
                # ボタン: KDL取込(430, 290)、直接取込(530, 290)、取込後(630, 290)、ログ確認(730, 290)
                # プロセスパネルのy座標は50なので、実際のy座標は140から（190-50=140）
                
                if ($i -eq 0) {
                    # 最初の行の場合のみ、V1抽出CSV格納元セクションを表示
                    # V1抽出CSV格納元ラベル
                    $v1CsvSourceLabel = New-Object System.Windows.Forms.Label
                    $v1CsvSourceLabel.Location = New-Object System.Drawing.Point(10, 30)
                    $v1CsvSourceLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvSourceLabel.Text = "V1抽出CSV格納元"
                    $v1CsvSourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 9, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvSourceLabel)
                    $script:v1CsvSourceLabel = $v1CsvSourceLabel
                    
                    # V1抽出CSV格納元パス入力
                    $v1CsvSourceTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvSourceTextBox.Location = New-Object System.Drawing.Point(10, 50)
                    $v1CsvSourceTextBox.Size = New-Object System.Drawing.Size(360, 30)
                    $v1CsvSourceTextBox.Text = "パス"
                    $v1CsvSourceTextBox.ReadOnly = $true
                    $v1CsvSourceTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvSourceTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvSourceTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvSourceTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvSourceTextBox.Add_Click({
                        if ($script:editMode) {
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "V1抽出CSV格納元フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $v1CsvSourceTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
                        }
                    })
                    $script:processPanel.Controls.Add($v1CsvSourceTextBox)
                    $script:v1CsvSourceTextBox = $v1CsvSourceTextBox
                }
                
                # プロセス行のレイアウト（行ごとに異なる）
                $x = 10
                $y = [int](140 + $row * 180)  # 行間隔を180pxに設定
                
                # タスク名
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(130, 30)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = $true
                $nameTextBox.BackColor = [System.Drawing.Color]::White
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $nameTextBox.Multiline = $false
                $nameTextBox.Height = 30
                $script:processPanel.Controls.Add($nameTextBox)
                
                # 1行目と2行目はKDL変換CSV格納元・格納先、V1抽出CSV格納先がある
                if ($i -lt 2) {
                    # KDL変換CSV格納元ラベル
                    $kdlSourceLabel = New-Object System.Windows.Forms.Label
                    $kdlSourceLabel.Location = New-Object System.Drawing.Point(150, [int]($y - 20))
                    $kdlSourceLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $kdlSourceLabel.Text = "KDL変換CSV格納元"
                    $kdlSourceLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($kdlSourceLabel)
                    
                    # KDL変換CSV格納元パス入力
                    $kdlSourceTextBox = New-Object System.Windows.Forms.TextBox
                    $kdlSourceTextBox.Location = New-Object System.Drawing.Point(150, $y)
                    $kdlSourceTextBox.Size = New-Object System.Drawing.Size(260, 30)
                    $kdlSourceTextBox.Text = "パス"
                    $kdlSourceTextBox.ReadOnly = $true
                    $kdlSourceTextBox.BackColor = [System.Drawing.Color]::White
                    $kdlSourceTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $kdlSourceTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $kdlSourceTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $kdlSourceTextBox.Add_Click({
                        if ($script:editMode) {
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "KDL変換CSV格納元フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $kdlSourceTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
                        }
                    })
                    $script:processPanel.Controls.Add($kdlSourceTextBox)
                    
                    # KDL変換CSV格納元の移動設定ボタン
                    $kdlSourceMoveButton = New-Object System.Windows.Forms.Button
                    $kdlSourceMoveButton.Location = New-Object System.Drawing.Point(415, $y)
                    $kdlSourceMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    $kdlSourceMoveButton.Text = "移動設定"
                    $kdlSourceMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc
                    $kdlSourceMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $kdlSourceMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    $kdlSourceMoveButton.FlatAppearance.BorderSize = 1
                    $kdlSourceMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $kdlSourceMoveButton.Visible = $script:editMode
                    $kdlSourceMoveButton.Tag = $i
                    $kdlSourceMoveButton.Add_Click({
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
                    $script:processPanel.Controls.Add($kdlSourceMoveButton)
                    
                    # KDL変換CSV格納先ラベル
                    $kdlDestLabel = New-Object System.Windows.Forms.Label
                    $kdlDestLabel.Location = New-Object System.Drawing.Point(490, [int]($y - 20))
                    $kdlDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $kdlDestLabel.Text = "KDL変換CSV格納先"
                    $kdlDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($kdlDestLabel)
                    
                    # KDL変換CSV格納先パス入力
                    $kdlDestTextBox = New-Object System.Windows.Forms.TextBox
                    $kdlDestTextBox.Location = New-Object System.Drawing.Point(490, $y)
                    $kdlDestTextBox.Size = New-Object System.Drawing.Size(230, 30)
                    $kdlDestTextBox.Text = "パス"
                    $kdlDestTextBox.ReadOnly = $true
                    $kdlDestTextBox.BackColor = [System.Drawing.Color]::White
                    $kdlDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $kdlDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $kdlDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $kdlDestTextBox.Add_Click({
                        if ($script:editMode) {
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "KDL変換CSV格納先フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $kdlDestTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
                        }
                    })
                    $script:processPanel.Controls.Add($kdlDestTextBox)
                    
                    # KDL変換CSV格納先の移動設定ボタン
                    $kdlDestMoveButton = New-Object System.Windows.Forms.Button
                    $kdlDestMoveButton.Location = New-Object System.Drawing.Point(725, $y)
                    $kdlDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    $kdlDestMoveButton.Text = "移動設定"
                    $kdlDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc
                    $kdlDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $kdlDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    $kdlDestMoveButton.FlatAppearance.BorderSize = 1
                    $kdlDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $kdlDestMoveButton.Visible = $script:editMode
                    $kdlDestMoveButton.Tag = $i
                    $kdlDestMoveButton.Add_Click({
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
                    $v1CsvDestTextBox.Add_Click({
                        if ($script:editMode) {
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "V1抽出CSV格納先フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $v1CsvDestTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
                        }
                    })
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # V1抽出CSV格納先の移動設定ボタン
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(725, [int]($y + 75))
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    $v1CsvDestMoveButton.Text = "移動設定"
                    $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $script:editMode
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
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
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    # ボタン行（KDL取込、直接取込、取込後、ログ確認）
                    $buttonY = [int]($y + 100)
                    
                    # KDL取込ボタン（赤色）
                    $kdlImportButton = New-Object System.Windows.Forms.Button
                    $kdlImportButton.Location = New-Object System.Drawing.Point(410, $buttonY)
                    $kdlImportButton.Size = New-Object System.Drawing.Size(90, 30)
                    $kdlImportButton.Text = "KDL取込"
                    $kdlImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 204)  # #ffcccc
                    $kdlImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $kdlImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(184, 84, 80)  # #b85450
                    $kdlImportButton.FlatAppearance.BorderSize = 1
                    $kdlImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $kdlImportButton.Enabled = $false  # 機能未実装
                    $script:processPanel.Controls.Add($kdlImportButton)
                    
                    # 直接取込ボタン（オレンジ）
                    $directImportButton = New-Object System.Windows.Forms.Button
                    $directImportButton.Location = New-Object System.Drawing.Point(510, $buttonY)
                    $directImportButton.Size = New-Object System.Drawing.Size(90, 30)
                    $directImportButton.Text = "直接取込"
                    $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)  # #ffe6cc
                    $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)  # #d79b00
                    $directImportButton.FlatAppearance.BorderSize = 1
                    $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $directImportButton.Enabled = $false  # 機能未実装
                    $script:processPanel.Controls.Add($directImportButton)
                    
                    # 取込後ボタン（オレンジ）
                    $afterImportButton = New-Object System.Windows.Forms.Button
                    $afterImportButton.Location = New-Object System.Drawing.Point(610, $buttonY)
                    $afterImportButton.Size = New-Object System.Drawing.Size(80, 30)
                    $afterImportButton.Text = "取込後"
                    $afterImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                    $afterImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $afterImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                    $afterImportButton.FlatAppearance.BorderSize = 1
                    $afterImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $afterImportButton.Enabled = $false  # 機能未実装
                    $script:processPanel.Controls.Add($afterImportButton)
                    
                    # ログ確認ボタン（緑）
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point(710, $buttonY)
                    $logButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $logButton.Text = "参照"
                    } else {
                        $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                    }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $processIdx = $i
                    $logButton.Add_Click({
                        Show-ProcessLog -ProcessIndex $processIdx
                    })
                    $script:processPanel.Controls.Add($logButton)
                    
                    # 4ページ目用のコントロール情報を保存（1行目・2行目）
                    $script:processControls += @{
                        NameTextBox = $nameTextBox
                        KdlSourceTextBox = $kdlSourceTextBox
                        KdlSourceMoveButton = $kdlSourceMoveButton
                        KdlDestTextBox = $kdlDestTextBox
                        KdlDestMoveButton = $kdlDestMoveButton
                        V1CsvDestTextBox = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        KdlImportButton = $kdlImportButton
                        DirectImportButton = $directImportButton
                        AfterImportButton = $afterImportButton
                        LogButton = $logButton
                    }
                } else {
                    # 3行目以降：V1抽出CSV格納先のみ
                    # V1抽出CSV格納先ラベル
                    $v1CsvDestLabel = New-Object System.Windows.Forms.Label
                    $v1CsvDestLabel.Location = New-Object System.Drawing.Point(150, [int]($y - 20))
                    $v1CsvDestLabel.Size = New-Object System.Drawing.Size(150, 20)
                    $v1CsvDestLabel.Text = "V1抽出CSV格納先"
                    $v1CsvDestLabel.Font = New-Object System.Drawing.Font("メイリオ", 8, [System.Drawing.FontStyle]::Bold)
                    $script:processPanel.Controls.Add($v1CsvDestLabel)
                    
                    # V1抽出CSV格納先パス入力
                    $v1CsvDestTextBox = New-Object System.Windows.Forms.TextBox
                    $v1CsvDestTextBox.Location = New-Object System.Drawing.Point(150, $y)
                    $v1CsvDestTextBox.Size = New-Object System.Drawing.Size(260, 30)
                    $v1CsvDestTextBox.Text = "パス"
                    $v1CsvDestTextBox.ReadOnly = $true
                    $v1CsvDestTextBox.BackColor = [System.Drawing.Color]::White
                    $v1CsvDestTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                    $v1CsvDestTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $v1CsvDestTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $v1CsvDestTextBox.Add_Click({
                        if ($script:editMode) {
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "V1抽出CSV格納先フォルダを選択してください"
                            $folderDialog.ShowNewFolderButton = $true
                            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                $v1CsvDestTextBox.Text = $folderDialog.SelectedPath
                            }
                            $folderDialog.Dispose()
                        }
                    })
                    $script:processPanel.Controls.Add($v1CsvDestTextBox)
                    
                    # V1抽出CSV格納先の移動設定ボタン
                    $v1CsvDestMoveButton = New-Object System.Windows.Forms.Button
                    $v1CsvDestMoveButton.Location = New-Object System.Drawing.Point(415, $y)
                    $v1CsvDestMoveButton.Size = New-Object System.Drawing.Size(60, 30)
                    $v1CsvDestMoveButton.Text = "移動設定"
                    $v1CsvDestMoveButton.BackColor = [System.Drawing.Color]::FromArgb(218, 232, 252)  # #dae8fc
                    $v1CsvDestMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $v1CsvDestMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 142, 191)  # #6c8ebf
                    $v1CsvDestMoveButton.FlatAppearance.BorderSize = 1
                    $v1CsvDestMoveButton.Font = New-Object System.Drawing.Font("メイリオ", 8)
                    $v1CsvDestMoveButton.Visible = $script:editMode
                    $v1CsvDestMoveButton.Tag = $i
                    $v1CsvDestMoveButton.Add_Click({
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
                    $script:processPanel.Controls.Add($v1CsvDestMoveButton)
                    
                    # ボタン行（直接取込、取込後、ログ確認）
                    $buttonY = [int]($y + 40)
                    
                    # 直接取込ボタン（オレンジ）
                    $directImportButton = New-Object System.Windows.Forms.Button
                    $directImportButton.Location = New-Object System.Drawing.Point(510, $buttonY)
                    $directImportButton.Size = New-Object System.Drawing.Size(90, 30)
                    $directImportButton.Text = "直接取込"
                    $directImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 204)  # #ffe6cc
                    $directImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $directImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 155, 0)  # #d79b00
                    $directImportButton.FlatAppearance.BorderSize = 1
                    $directImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $directImportButton.Enabled = $false  # 機能未実装
                    $script:processPanel.Controls.Add($directImportButton)
                    
                    # 取込後ボタン（オレンジ）
                    $afterImportButton = New-Object System.Windows.Forms.Button
                    $afterImportButton.Location = New-Object System.Drawing.Point(610, $buttonY)
                    $afterImportButton.Size = New-Object System.Drawing.Size(80, 30)
                    $afterImportButton.Text = "取込後"
                    $afterImportButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 153)  # #ffcc99
                    $afterImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $afterImportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(214, 182, 86)  # #d6b656
                    $afterImportButton.FlatAppearance.BorderSize = 1
                    $afterImportButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $afterImportButton.Enabled = $false  # 機能未実装
                    $script:processPanel.Controls.Add($afterImportButton)
                    
                    # ログ確認ボタン（緑）
                    $logButton = New-Object System.Windows.Forms.Button
                    $logButton.Location = New-Object System.Drawing.Point(710, $buttonY)
                    $logButton.Size = New-Object System.Drawing.Size(80, 30)
                    if ($script:editMode) {
                        $logButton.Text = "参照"
                    } else {
                        $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                    }
                    $logButton.BackColor = [System.Drawing.Color]::FromArgb(213, 232, 212)  # #d5e8d4
                    $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(130, 179, 102)  # #82b366
                    $logButton.FlatAppearance.BorderSize = 1
                    $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                    $processIdx = $i
                    $logButton.Add_Click({
                        Show-ProcessLog -ProcessIndex $processIdx
                    })
                    $script:processPanel.Controls.Add($logButton)
                    
                    # 4ページ目用のコントロール情報を保存（3行目以降）
                    $script:processControls += @{
                        NameTextBox = $nameTextBox
                        V1CsvDestTextBox = $v1CsvDestTextBox
                        V1CsvDestMoveButton = $v1CsvDestMoveButton
                        DirectImportButton = $directImportButton
                        AfterImportButton = $afterImportButton
                        LogButton = $logButton
                    }
                }
            } else {
                # 5ページ目以降：従来のレイアウト
                $x = [int](10 + $col * 440)
                $y = [int](10 + $row * 60)
                
                # テキストボックス（タスク名表示用）
                $nameTextBox = New-Object System.Windows.Forms.TextBox
                $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
                $nameTextBox.Size = New-Object System.Drawing.Size(140, 40)
                $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
                $nameTextBox.ReadOnly = $true
                $nameTextBox.BackColor = [System.Drawing.Color]::White
                $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $nameTextBox.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $nameTextBox.Multiline = $true
                $nameTextBox.Height = 40
                $script:processPanel.Controls.Add($nameTextBox)
                
                # ファイル移動設定ボタン（水色）- 編集モードONの時のみ表示
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
                } else {
                    $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "実行" }
                }
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
                $logX = [int]($x + 330)
                $logButton.Location = New-Object System.Drawing.Point($logX, $y)
                $logButton.Size = New-Object System.Drawing.Size(80, 40)
                if ($script:editMode) {
                    $logButton.Text = "参照"
                } else {
                    $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "ログ確認" }
                }
                $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
                $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                $logButton.FlatAppearance.BorderSize = 1
                $logButton.Font = New-Object System.Drawing.Font("メイリオ", 9)
                $processIdx = $i
                $logButton.Add_Click({
                    Show-ProcessLog -ProcessIndex $processIdx
                })
                $script:processPanel.Controls.Add($logButton)
                
                # 5ページ目以降用のコントロール情報を保存
                $script:processControls += @{
                    NameTextBox = $nameTextBox
                    FileMoveButton = $fileMoveButton
                    ExecuteButton = $executeButton
                    LogButton = $logButton
                }
            }
        }
    }
    
    # ページ情報の更新
    $script:pageLabel.Text = "ページ $($script:currentPage + 1) / $totalPages"
    
    # タイトルの更新（ページJSONから読み込む）
    $pageTitle = ""
    $pageConfig = $script:pages[$script:currentPage]
    if ($pageConfig.JsonPath) {
        $pageJsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        } else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
        if (Test-Path $pageJsonPath) {
            try {
                $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
                if ($pageJson.Title) {
                    $pageTitle = $pageJson.Title
                }
            } catch {
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
    for ($i = 0; $i -lt $script:processControls.Count; $i++) {
        $ctrlGroup = $script:processControls[$i]
        if ($ctrlGroup -and $ctrlGroup.FileMoveButton) {
            if ($isPage1) {
                # 1ページ目：常に表示、テキストを編集モードに応じて更新（ONの時は「参照」、OFFの時は「チェック」）
                $ctrlGroup.FileMoveButton.Visible = $true
                if ($script:editMode) {
                    $ctrlGroup.FileMoveButton.Text = "参照"
                } else {
                    $ctrlGroup.FileMoveButton.Text = "チェック"  # 設計書通り「チェック」と表示
                }
            } elseif ($isPage2) {
                # 2ページ目：常に表示、テキストを編集モードに応じて更新（ONの時は「参照」、OFFの時は「セット」）
                $ctrlGroup.FileMoveButton.Visible = $true
                if ($script:editMode) {
                    $ctrlGroup.FileMoveButton.Text = "参照"
                } else {
                    $ctrlGroup.FileMoveButton.Text = "セット"
                }
            } else {
                # その他のページ：編集モードONの時のみ表示
                $ctrlGroup.FileMoveButton.Visible = $script:editMode
            }
        }
    }
    
    # ページパスの読み込み
    Load-PagePaths
    
    # ページに応じてレイアウトを調整
    if ($useDrawioLayout) {
        # 1ページ目・2ページ目：ファイル移動セクションを非表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # ログ格納セクションの位置を調整（370px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 370)
        }
        
        # ログ格納ボタンの位置を調整（2ページ目は右側に配置）
        if ($script:logStorageButton) {
            if ($isPage2) {
                # 2ページ目：右側に配置（700px x座標）
                $script:logStorageButton.Location = New-Object System.Drawing.Point(690, 35)
            } else {
                # 1ページ目：左側に配置（340px x座標）
                $script:logStorageButton.Location = New-Object System.Drawing.Point(340, 35)
            }
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
    } elseif ($isPage3) {
        # 3ページ目：JAVA移行ツール実行のレイアウト
        # ヘッダーの背景色を緑色に変更
        if ($script:headerPanel) {
            $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(147, 196, 125)  # #93C47D
        }
        
        # ファイル移動セクションを非表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # ログ格納セクションの位置を調整（370px y座標）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Location = New-Object System.Drawing.Point(0, 370)
        }
        
        # ログ格納ボタンの位置を調整（400px x座標）
        if ($script:logStorageButton) {
            $script:logStorageButton.Location = New-Object System.Drawing.Point(390, 35)
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
        
        # プロセスパネルの高さを調整（370px）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 370)
        }
    } elseif ($isPage4) {
        # 4ページ目：SQLLOADER実行のレイアウト
        # ヘッダーの背景色を青色に変更
        if ($script:headerPanel) {
            $script:headerPanel.BackColor = [System.Drawing.Color]::FromArgb(27, 161, 226)  # #1ba1e2
        }
        
        # ファイル移動セクションを非表示
        if ($script:fileMovePanel) {
            $script:fileMovePanel.Visible = $false
        }
        
        # ログ格納セクションを非表示（4ページ目にはログ格納セクションがない）
        if ($script:logStoragePanel) {
            $script:logStoragePanel.Visible = $false
        }
        
        # ログ出力エリアの位置を調整（690px y座標、790px幅、150px高さ）
        if ($script:logTextBox) {
            $script:logTextBox.Location = New-Object System.Drawing.Point(10, 690)
            $script:logTextBox.Size = New-Object System.Drawing.Size(790, 150)
        }
        
        # フォームの高さを調整（860px）
        if ($script:form) {
            $script:form.Size = New-Object System.Drawing.Size(900, 860)
        }
        
        # プロセスパネルの高さを調整（680px）
        if ($script:processPanel) {
            $script:processPanel.Size = New-Object System.Drawing.Size(900, 680)
        }
    } else {
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
