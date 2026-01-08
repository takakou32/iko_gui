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
    
    # 既存のコントロールをクリア
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup) {
            $script:processPanel.Controls.Remove($ctrlGroup.NameTextBox)
            $script:processPanel.Controls.Remove($ctrlGroup.FileMoveButton)
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
            
            # コントロールの位置計算（列幅を440に拡大）
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
            # ボタンのTagプロパティにインデックスを保存（クロージャーの問題を回避）
            $fileMoveButton.Tag = $i
            $fileMoveButton.Add_Click({
                # クリック時にボタンのTagからインデックスを取得
                $clickedProcessIdx = $this.Tag
                # processControls配列から該当するNameTextBoxを取得
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
            $processIdx = $i  # クロージャーの問題を回避するため、ローカル変数に保存
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
            $processIdx = $i  # クロージャーの問題を回避するため、ローカル変数に保存
            $logButton.Add_Click({
                Show-ProcessLog -ProcessIndex $processIdx
            })
            $script:processPanel.Controls.Add($logButton)
            
            $script:processControls += @{
                NameTextBox = $nameTextBox
                FileMoveButton = $fileMoveButton
                ExecuteButton = $executeButton
                LogButton = $logButton
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
    
    # 移動設定ボタンの表示/非表示を編集モードに応じて更新
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup -and $ctrlGroup.FileMoveButton) {
            $ctrlGroup.FileMoveButton.Visible = $script:editMode
        }
    }
    
    # ページパスの読み込み
    Load-PagePaths
}
