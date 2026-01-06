# PowerShellスクリプト - GUIアプリケーション
# エンコーディング: UTF-8 BOM付
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
$script:pageProcessCache = @()
$script:editMode = $false

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
    
    # 移行データファイル移動元
    $sourcePath = if ($pageConfig.SourcePath) { $pageConfig.SourcePath } else { "" }
    if ($sourcePath -and $sourcePath -ne "パス" -and $sourcePath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($sourcePath)) {
                $sourcePath = Join-Path $PSScriptRoot $sourcePath
            }
            $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
            $script:sourcePathTextBox.Text = $sourcePath
        } catch {
            $script:sourcePathTextBox.Text = if ($pageConfig.SourcePath) { $pageConfig.SourcePath } else { "パス" }
        }
    } else {
        $script:sourcePathTextBox.Text = "パス"
    }
    
    # 移行データファイル移動先
    $destPath = if ($pageConfig.DestinationPath) { $pageConfig.DestinationPath } else { "" }
    if ($destPath -and $destPath -ne "パス" -and $destPath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($destPath)) {
                $destPath = Join-Path $PSScriptRoot $destPath
            }
            $destPath = [System.IO.Path]::GetFullPath($destPath)
            $script:destPathTextBox.Text = $destPath
        } catch {
            $script:destPathTextBox.Text = if ($pageConfig.DestinationPath) { $pageConfig.DestinationPath } else { "パス" }
        }
    } else {
        $script:destPathTextBox.Text = "パス"
    }
    
    # ログ格納先
    $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
    if ($logStoragePath -and $logStoragePath -ne "パス" -and $logStoragePath -ne "") {
        # 相対パスの場合は絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($logStoragePath)) {
                $logStoragePath = Join-Path $PSScriptRoot $logStoragePath
            }
            $logStoragePath = [System.IO.Path]::GetFullPath($logStoragePath)
            $script:logStoragePathTextBox.Text = $logStoragePath
        } catch {
            $script:logStoragePathTextBox.Text = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "パス" }
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
    
    # config.jsonを再読み込み
    if (-not (Test-Path $configPath)) {
        Write-Log "設定ファイルが見つかりません: $configPath" "ERROR"
        return $false
    }
    
    try {
        $configContent = Get-Content $configPath -Encoding UTF8 -Raw | ConvertFrom-Json
        
        if (-not $configContent.Pages -or $script:currentPage -ge $configContent.Pages.Count) {
            Write-Log "ページインデックスが範囲外です" "ERROR"
            return $false
        }
        
        $pageConfig = $configContent.Pages[$script:currentPage]
        
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
            $pageConfig.SourcePath = $relativeSourcePath
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
            $pageConfig.DestinationPath = $relativeDestPath
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
            $pageConfig.LogStoragePath = $relativeLogPath
        }
        
        # config.jsonに保存
        $configContent | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
        
        # メモリ上のpagesも更新
        $script:pages[$script:currentPage].SourcePath = if ($SourcePath) { $pageConfig.SourcePath } else { $script:pages[$script:currentPage].SourcePath }
        $script:pages[$script:currentPage].DestinationPath = if ($DestinationPath) { $pageConfig.DestinationPath } else { $script:pages[$script:currentPage].DestinationPath }
        $script:pages[$script:currentPage].LogStoragePath = if ($LogStoragePath) { $pageConfig.LogStoragePath } else { $script:pages[$script:currentPage].LogStoragePath }
        
        Write-Log "ページパスを保存しました" "INFO"
        return $true
    } catch {
        Write-Log "設定ファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR"
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
            if (Test-Path $logDir) {
                $folderDialog.SelectedPath = $logDir
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
            $logX = [int]($x + 300)
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
    
    # ページパスの読み込み
    Load-PagePaths
}

# GUIフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "プロセス実行GUI"
$form.Size = New-Object System.Drawing.Size(800, 900)
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
$processPanel.Size = New-Object System.Drawing.Size(800, 280)
$processPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($processPanel)
$script:processPanel = $processPanel

# ファイル移動セクション（黄色/ベージュ背景）
$fileMovePanel = New-Object System.Windows.Forms.Panel
$fileMovePanel.Location = New-Object System.Drawing.Point(0, 330)
$fileMovePanel.Size = New-Object System.Drawing.Size(800, 120)
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
$logStoragePanel.Location = New-Object System.Drawing.Point(0, 450)
$logStoragePanel.Size = New-Object System.Drawing.Size(800, 80)
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
$logOutputLabel.Location = New-Object System.Drawing.Point(10, 540)
$logOutputLabel.Size = New-Object System.Drawing.Size(200, 20)
$logOutputLabel.Text = "ログ出力(全ファイル)"
$logOutputLabel.Font = New-Object System.Drawing.Font("メイリオ", 9)
$form.Controls.Add($logOutputLabel)

# ログ表示エリア
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 565)
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

# 初期ページパスの読み込み（Update-ProcessControls内でLoad-PagePathsが呼ばれるが、念のため）
Load-PagePaths

# 初期メッセージ
Write-Log "アプリケーションを起動しました" "INFO"
Write-Log "設定ファイル: $configPath" "INFO"
Write-Log "ページ数: $($script:pages.Count)" "INFO"

# フォームを表示
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
