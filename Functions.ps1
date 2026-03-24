# PowerShellスクリプト - 関数定義
# エンコーディング: UTF-8 BOM付

# 現在のページのプロセス一覧を取得
# 現在のページのプロセス一覧を取得
# モダンなフォルダー選択ダイアログを使用するためのクラス定義
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class FolderSelectDialog
{
    [DllImport("shell32.dll")]
    private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, ref Guid riid, out IShellItem ppv);

    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    private const string IID_IShellItem = "43826d1e-e718-42ee-bc55-a1e261c37bfe";
    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;

    public string InitialDirectory { get; set; }
    public string Title { get; set; }

    public bool ShowDialog(out string selectedPath)
    {
        selectedPath = null;
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialog();
        
        try
        {
            dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
            
            if (!string.IsNullOrEmpty(Title))
            {
                dialog.SetTitle(Title);
            }

            if (!string.IsNullOrEmpty(InitialDirectory))
            {
                IShellItem item;
                Guid riid = new Guid(IID_IShellItem);
                if (SHCreateItemFromParsingName(InitialDirectory, IntPtr.Zero, ref riid, out item) == 0)
                {
                    dialog.SetFolder(item);
                }
            }

            if (dialog.Show(GetActiveWindow()) == 0) // S_OK
            {
                IShellItem result;
                dialog.GetResult(out result);
                string path;
                result.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out path);
                selectedPath = path;
                return true;
            }
        }
        catch (Exception)
        {
            // Fallback needed or just return false
        }
        finally
        {
            Marshal.ReleaseComObject(dialog);
        }

        return false;
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    [ClassInterface(ClassInterfaceType.None)]
    private class FileOpenDialog { }

    [ComImport]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(); // Placeholder
        void SetFileTypeIndex(); // Placeholder
        void GetFileTypeIndex(); // Placeholder
        void Advise(); // Placeholder
        void Unadvise(); // Placeholder
        void SetOptions(uint fos);
        void GetOptions(); // Placeholder
        void SetDefaultFolder(); // Placeholder
        void SetFolder(IShellItem psi);
        void GetFolder(); // Placeholder
        void GetCurrentSelection(); // Placeholder
        void SetFileName(); // Placeholder
        void GetFileName(); // Placeholder
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel(); // Placeholder
        void SetFileNameLabel(); // Placeholder
        void GetResult(out IShellItem ppsi);
        void AddPlace(); // Placeholder
        void SetDefaultExtension(); // Placeholder
        void Close(); // Placeholder
        void SetClientGuid(); // Placeholder
        void ClearClientData(); // Placeholder
        void SetFilter(); // Placeholder
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(); // Placeholder
        void GetParent(); // Placeholder
        void GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(); // Placeholder
        void Compare(); // Placeholder
    }

    private enum SIGDN : uint
    {
        SIGDN_FILESYSPATH = 0x80058000
    }
}
"@

function Show-FolderBrowser {
    param (
        [string]$InitialDirectory,
        [string]$Description = "フォルダーを選択してください"
    )

    $dialog = New-Object FolderSelectDialog
    
    if ($InitialDirectory) {
        # Check if we should resolve relative path (for log settings)
        if ($Description -like "*ログ*") {
            try {
                # Try to resolve relative path to global log path
                if (-not [System.IO.Path]::IsPathRooted($InitialDirectory)) {
                    $InitialDirectory = Join-Path $script:globalLogPath $InitialDirectory
                }
            }
            catch {
                # Ignore path errors
            }
        }
    }
    
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.Title = $Description
    
    [string]$path = $null
    if ($dialog.ShowDialog([ref]$path)) {
        return $path
    }
    return $null
}

# パスをエクスプローラで開く（フォルダはそのまま開く、ファイルは親フォルダを開いてファイルを選択＝実行しない）
function Open-PathInExplorer {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) { return }
    if (Test-Path $Path -PathType Container) {
        Start-Process explorer -ArgumentList "`"$Path`""
    }
    else {
        Start-Process explorer -ArgumentList "/select,`"$Path`""
    }
}

# ログパス解決関数
function Resolve-LogPath {
    param([string]$SubPath)
    
    if (-not $SubPath) {
        return $script:globalLogPath
    }
    
    if ([System.IO.Path]::IsPathRooted($SubPath)) {
        return $SubPath
    }
    
    return Join-Path $script:globalLogPath $SubPath
}

function Get-CurrentPageProcesses {
    if ($script:currentPage -ge $script:pages.Count) {
        return @()
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # JsonPathが指定されている場合は、そのJSONファイルを読み込む
    if ($pageConfig.JsonPath) {
        $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        }
        else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
        
        if (Test-Path $jsonPath) {
            try {
                $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                if ($pageJson.Processes) {
                    return @($pageJson.Processes)
                }
                else {
                    Write-Log "JSONファイルにProcessesが含まれていません: $jsonPath" "WARN"
                    return @()
                }
            }
            catch {
                Write-Log "JSONファイルの読み込みに失敗しました: $jsonPath - $($_.Exception.Message)" "ERROR"
                return @()
            }
        }
        else {
            Write-Log "JSONファイルが見つかりません: $jsonPath" "ERROR"
            return @()
        }
    }
    
    # JsonPathが指定されていない場合は、直接Processesを使用（後方互換性）
    if ($pageConfig.Processes) {
        return @($pageConfig.Processes)
    }
    
    return @()
}

# ログ出力関数
function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [int]$ProcessIndex = -1, [string]$LogDir = $null)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # 単一セッションログファイルへの出力（ユーザー要望により一元化）
    if ($script:sessionLogFile) {
        # プロセスインデックスがある場合は識別子を追加
        $logPrefix = ""
        if ($ProcessIndex -ge 0) {
            $logPrefix = "[Page:$($script:currentPage + 1) Process:$($ProcessIndex + 1)] "
        }
        
        $fileLogMessage = "[$timestamp] [$Level] $logPrefix$Message"
        
        # ログファイルへの追記（ロック競合を避けるため簡易的なリトライを入れるか、あるいは単一スレッド前提とする）
        # GUIイベントハンドラ内であればメインスレッドなので競合はしにくいが、念のため
        try {
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::AppendAllText($script:sessionLogFile, $fileLogMessage + "`r`n", $utf8NoBom)
        }
        catch {
            # ログ出力失敗時はコンソールに出すくらいしかできない
            Write-Host "Log Write Failed: $_"
        }
    }
    
    # GUIのログ表示エリアに追加
    $script:logTextBox.AppendText("$logMessage`r`n")
    $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length
    $script:logTextBox.ScrollToCaret()
    
    Write-Host $logMessage
}

# Batファイル実行関数
# Batファイル実行関数
function Invoke-BatchFile {
    param(
        [string]$BatchPath,
        [string]$DisplayName,
        [int]$ProcessIndex,
        [string[]]$Arguments = @()
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
    }
    catch {
        Write-Log "バッチファイルパスの正規化に失敗しました: $BatchPath - $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
    
    if (-not (Test-Path $BatchPath)) {
        Write-Log "バッチファイルが見つかりません: $BatchPath" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("バッチファイルが見つかりません。`n$BatchPath", "実行エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    
    Write-Log "バッチファイル実行開始: $DisplayName ($BatchPath)" "INFO" $ProcessIndex
    
    # 実行ディレクトリ（バッチファイルのある場所）
    $workingDir = Split-Path -Parent $BatchPath
    
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$BatchPath`" $Arguments"
        $psi.WorkingDirectory = $workingDir
        $psi.UseShellExecute = $true
        # ウィンドウを表示するため、リダイレクトを無効化（ユーザー要望）
        # $psi.RedirectStandardOutput = $true
        # $psi.RedirectStandardError = $true
        # $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $false
        $psi.RedirectStandardError = $false
        $psi.CreateNoWindow = $false
        
        # エンコーディング設定（Shift-JIS）
        if ($psi.RedirectStandardOutput) {
            $psi.StandardOutputEncoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        }
        if ($psi.RedirectStandardError) {
            $psi.StandardErrorEncoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        }
        
        # $process = New-Object System.Diagnostics.Process
        # $process.StartInfo = $psi
        # $process.EnableRaisingEvents = $true
        
        # 出力ハンドラ（Write-Logにリダイレクト）
        $outputHandler = {
            param($sender, $e)
            if ($e.Data) {
                # メインスレッドでの実行を確実にするため、Write-Logを直接呼び出す
                # 注意: イベントハンドラは別スレッドで実行される可能性があるが、
                # Write-Log内でファイル書き込みを行っているため（排他制御は簡易的だが）、
                # ここではそのまま呼び出す。UI更新系はInvokeが必要かもしれないが、
                # Write-Log内のUI更新はテキストボックスへの追記なので、
                # cross-thread operation エラーが出る可能性がある。
                # そこで、フォームのInvokeを使用する。
                $script:mainForm.Invoke([action] { Write-Log $e.Data "INFO" $ProcessIndex })
            }
        }
        
        $errorHandler = {
            param($sender, $e)
            if ($e.Data) {
                $script:mainForm.Invoke([action] { Write-Log $e.Data "ERROR" $ProcessIndex })
            }
        }
        
        # スタティックメソッドでプロセスを開始（0引数エラー対策）
        $process = [System.Diagnostics.Process]::Start($psi)
        if ($psi.RedirectStandardOutput) {
            $process.BeginOutputReadLine()
        }
        if ($psi.RedirectStandardError) {
            $process.BeginErrorReadLine()
        }
        
        # UIの応答性を維持しながら待機
        while (-not $process.HasExited) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 50
        }
        
        $exitCode = $process.ExitCode
        $process.Dispose()
        
        if ($exitCode -eq 0) {
            Write-Log "バッチファイル実行完了 (ExitCode: 0)" "INFO" $ProcessIndex
            return $true
        }
        elseif ($exitCode -eq -1073741510) {
            # 0xC000013A (STATUS_CONTROL_C_EXIT): ユーザーによるウィンドウ「×」ボタンクローズ
            Write-Log "バッチファイル実行完了 (ウィンドウ切断: $exitCode)" "INFO" $ProcessIndex
            return $true
        }
        elseif ($exitCode -eq 3) {
            # ExitCode 3: ユーザー申告による正常終了値
            Write-Log "バッチファイル実行完了 (ExitCode: 3)" "INFO" $ProcessIndex
            return $true
        }
        else {
            Write-Log "バッチファイル実行エラー (ExitCode: $exitCode)" "ERROR" $ProcessIndex
            [System.Windows.Forms.MessageBox]::Show("バッチファイルの実行中にエラーが発生しました。`nログを確認してください。", "実行エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    }
    catch {
        Write-Log "実行例外: $($_.Exception.Message)" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("実行中に例外が発生しました。`n$($_.Exception.Message)", "例外", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# ログ格納用バッチファイルパス保存関数
function Save-LogStorageBatchFile {
    param([string]$BatchFilePath)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN"
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $jsonPath)) {
        Write-Log "JSONファイルが見つかりません: $jsonPath" "ERROR"
        return $false
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # LogStorageBatchFile要素が存在しない場合は作成
        if (-not $jsonContent.LogStorageBatchFile) {
            $jsonContent | Add-Member -MemberType NoteProperty -Name "LogStorageBatchFile" -Value @{
                Name = "ログ格納用バッチファイル"
                Path = ""
            }
        }
        
        # 相対パスに変換（可能な場合）
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            # ユーザー要望により、GlobalLogPath（ツール格納場所）を基準とする
            $basePath = if ($script:globalLogPath) { 
                [System.IO.Path]::GetFullPath($script:globalLogPath).TrimEnd('\', '/') 
            }
            else { 
                [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/') 
            }
            
            $targetPath = [System.IO.Path]::GetFullPath($BatchFilePath).TrimEnd('\', '/')
            
            # デバッグ用ログ
            Write-Log "Path Check - Base: '$basePath', Target: '$targetPath'" "INFO"
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                if ([string]::IsNullOrEmpty($relative)) {
                    $relative = Split-Path $targetPath -Leaf
                }
                $relative
            }
            else {
                # 基準パス外の場合はエラーとして処理
                Write-Log "バッチファイルは共通パス（$basePath）配下に配置する必要があります: $targetPath" "ERROR"
                [void][System.Windows.Forms.MessageBox]::Show("バッチファイルは共通パス（ツール格納場所）配下に配置する必要があります。`n共通パス: $basePath`n選択されたパス: $targetPath", "設定エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return $false
            }
        }
        catch {
            Write-Log "パス変換エラー: $($_.Exception.Message)" "ERROR"
            $BatchFilePath
        }
        
        $jsonContent.LogStorageBatchFile.Path = $relativePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "ログ格納用バッチファイルパスを保存しました: $relativePath" "INFO"
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# バッチファイルパス保存関数
function Save-BatchFilePath {
    param([int]$ProcessIndex, [string]$BatchFilePath, [int]$BatchIndex = 0)
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # LogStoragePath（共通パス）を取得
    # ユーザー要望により、GlobalLogPath（ツール格納場所）を使用する
    $logStoragePath = if ($script:globalLogPath) { $script:globalLogPath } else { "" }
    
    # GlobalLogPathが設定されていない場合は、ページ設定のLogStoragePathを使用（後方互換性）
    if (-not $logStoragePath) {
        $logStoragePath = if ($pageConfig.LogStoragePath) { $pageConfig.LogStoragePath } else { "" }
        
        # メモリ上にない場合、JSONファイルから直接読み込む試み
        if (-not $logStoragePath -and $pageConfig.JsonPath) {
            $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
                $pageConfig.JsonPath
            }
            else {
                $path1 = Join-Path $PSScriptRoot $pageConfig.JsonPath
                if (Test-Path $path1) {
                    $path1
                }
                elseif ($script:configDir -and (Test-Path (Join-Path $script:configDir $pageConfig.JsonPath))) {
                    Join-Path $script:configDir $pageConfig.JsonPath
                }
                else {
                    $path1
                }
            }
            
            if (Test-Path $jsonPath) {
                try {
                    $pageJson = Get-Content $jsonPath -Encoding UTF8 | ConvertFrom-Json
                    if ($pageJson.LogStoragePath) {
                        $logStoragePath = $pageJson.LogStoragePath
                        # メモリ上の設定も更新（次回以降のためにキャッシュ）
                        $pageConfig.LogStoragePath = $logStoragePath
                    }
                }
                catch {}
            }
        }
    }
    
    # 相対パスの場合は絶対パスに展開（比較用）
    $fullBatchPath = [System.IO.Path]::GetFullPath($BatchFilePath)
    
    # LogStoragePathが設定されている場合、パス制限チェック
    if ($logStoragePath) {
        $fullLogStoragePath = [System.IO.Path]::GetFullPath($logStoragePath)
        
        # 大文字小文字を区別せずに比較（共通パス配下かチェック）
        if (-not $fullBatchPath.StartsWith($fullLogStoragePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log "バッチファイルは共通パス（$fullLogStoragePath）配下に配置する必要があります: $fullBatchPath" "ERROR" $ProcessIndex
            [System.Windows.Forms.MessageBox]::Show("バッチファイルは共通パス（LogStoragePath）配下に配置する必要があります。`n共通パス: $fullLogStoragePath", "設定エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
        
        # 相対パスに変換
        $relativePath = $fullBatchPath.Substring($fullLogStoragePath.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
        $BatchFilePath = $relativePath
    }
    
    # ページ設定の保存処理
    # 既存のロジック（JSON保存）を維持しつつ、BatchFiles配列の更新を行う
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        # 指定されたインデックスまで配列を拡張
        while ($process.BatchFiles.Count -le $BatchIndex) {
            $process.BatchFiles += @{ Name = "バッチファイル"; Path = "" }
        }
        
        # パスを更新（相対パス）
        $process.BatchFiles[$BatchIndex].Path = $BatchFilePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        
        # メモリ上の設定も更新（重要：UI更新のため）
        if ($script:pages[$script:currentPage].Processes) {
            # メモリ上のProcesses配列も更新する必要があるが、
            # ここでは簡易的にJSON再読み込みを促すか、あるいは直接メモリを更新する
            # Functions.ps1の他の箇所ではLoad-PageConfigなどで再読み込みしている可能性がある
            # ここではJSON保存成功を返すのみとする（呼び出し元でUpdate-ProcessControlsなどが呼ばれるため再描画時に再読み込みされることを期待）
            # しかし、メモリ上のキャッシュが更新されないと即座に反映されない場合がある
            # 安全のため、メモリ上のオブジェクトも更新する
            try {
                if ($script:pages[$script:currentPage].Processes[$ProcessIndex].BatchFiles) {
                    # 配列拡張が必要な場合もあるが、複雑になるため、ここではJSON保存を主とする
                    # Reload-PageConfigがあれば呼びたいところ
                }
            }
            catch {}
        }

        Write-Log "バッチファイルパスを保存しました (Index: $BatchIndex): $BatchFilePath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "バッチファイルパスの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# プロセス名保存関数
function Save-ProcessName {
    param(
        [int]$ProcessIndex,
        [string]$ProcessName
    )
    
    if ($script:currentPage -ge $script:pages.Count) {
        Write-Log "ページインデックスが範囲外です" "ERROR" $ProcessIndex
        return $false
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        $process.Name = $ProcessName
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "プロセス名を保存しました: $ProcessName" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}




# プロセスDestinationPath保存関数（ページ3用）
function Save-ProcessDestinationPath {
    param([int]$ProcessIndex, [string]$DestinationPath)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        # DestinationPathプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name "DestinationPath" -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name "DestinationPath" -Value ""
        }
        
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            $basePath = Get-CommonBasePath
            $targetPath = [System.IO.Path]::GetFullPath($DestinationPath).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                # 相対パスが空文字列の場合（選択パスが共通基準パスと完全に同じ場合）は絶対パスをそのまま保存するのではなく、"."とするか、空にするか
                # 既存ロジックに合わせてそのまま返す
                $relative
            }
            else {
                $DestinationPath
            }
        }
        catch {
            $DestinationPath
        }
        
        $process.DestinationPath = $relativePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "プロセスDestinationPathを保存しました: $relativePath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}



# 共通基準パス取得関数
# 共通基準パス取得関数
function Get-CommonBasePath {
    # ユーザー要望により、GlobalLogPathを最優先
    if ($script:globalLogPath) {
        return [System.IO.Path]::GetFullPath($script:globalLogPath).TrimEnd('\', '/')
    }
    
    # 次にページ固有のLogStoragePath
    $pageConfig = $script:pages[$script:currentPage]
    if ($pageConfig.LogStoragePath) {
        $path = $pageConfig.LogStoragePath
        if (-not [System.IO.Path]::IsPathRooted($path)) {
            $path = Join-Path $PSScriptRoot $path
        }
        return [System.IO.Path]::GetFullPath($path).TrimEnd('\', '/')
    }
    
    # フォールバック
    return [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
}

# プロセスKDL変換CSV格納先パス保存関数（ページ4用）
function Save-ProcessKdlDestPath {
    param([int]$ProcessIndex, [string]$KdlDestPath)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        # KdlDestPathプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name "KdlDestPath" -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name "KdlDestPath" -Value ""
        }
        
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            $basePath = Get-CommonBasePath
            $targetPath = [System.IO.Path]::GetFullPath($KdlDestPath).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                # 相対パスが空文字列の場合（選択パスが基準パスと完全に同じ場合）は絶対パスをそのまま保存するのではなく、"."とするか、空にするか
                # ここではディレクトリ指定なので "" (空) になると困るかもしれないが、Join-Pathで空文字は無視されるのでOK
                # ただし、JSON上 "" だと未設定と区別がつかない可能性があるが、現状のロジックでは "" はパスとして認識される
                $relative
            }
            else {
                $KdlDestPath
            }
        }
        catch {
            $KdlDestPath
        }
        
        $process.KdlDestPath = $relativePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "プロセスKDL変換CSV格納先パスを保存しました: $relativePath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# プロセスKDL変換CSV格納元パス保存関数（ページ4用）
function Save-ProcessKdlSourcePath {
    param([int]$ProcessIndex, [string]$KdlSourcePath)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        # KdlSourcePathプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name "KdlSourcePath" -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name "KdlSourcePath" -Value ""
        }
        
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            $basePath = Get-CommonBasePath
            $targetPath = [System.IO.Path]::GetFullPath($KdlSourcePath).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                $relative
            }
            else {
                $KdlSourcePath
            }
        }
        catch {
            $KdlSourcePath
        }
        
        $process.KdlSourcePath = $relativePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "プロセスKDL変換CSV格納元パスを保存しました: $relativePath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# プロセスV1抽出CSV格納先パス保存関数（ページ4用）
function Save-ProcessV1CsvDestPath {
    param([int]$ProcessIndex, [string]$V1CsvDestPath)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        # V1CsvDestPathプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name "V1CsvDestPath" -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name "V1CsvDestPath" -Value ""
        }
        
        # 相対パスに変換（可能な場合）
        $relativePath = try {
            $basePath = Get-CommonBasePath
            $targetPath = [System.IO.Path]::GetFullPath($V1CsvDestPath).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                $relative
            }
            else {
                $V1CsvDestPath
            }
        }
        catch {
            $V1CsvDestPath
        }
        
        $process.V1CsvDestPath = $relativePath
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "プロセスV1抽出CSV格納先パスを保存しました: $relativePath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ログ出力フォルダパス保存関数
function Save-ProcessLogOutputDir {
    param(
        [int]$ProcessIndex, 
        [string]$LogOutputDir,
        [int]$LogIndex = 1
    )
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        $propName = if ($LogIndex -eq 2) { "LogOutputDir2" } else { "LogOutputDir" }
        
        # LogOutputDirプロパティが存在しない場合は追加
        if (-not (Get-Member -InputObject $process -Name $propName -MemberType NoteProperty)) {
            Add-Member -InputObject $process -MemberType NoteProperty -Name $propName -Value $LogOutputDir
        }
        else {
            $process.$propName = $LogOutputDir
        }
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "ログ出力フォルダパス($propName)を保存しました: $LogOutputDir" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ページパス読み込み関数
function Update-PagePaths {
    if ($script:currentPage -ge $script:pages.Count) {
        return
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # ページJSONファイルから設定を読み込む
    $pageJsonPath = $null
    if ($pageConfig.JsonPath) {
        $pageJsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        }
        else {
            $path1 = Join-Path $PSScriptRoot $pageConfig.JsonPath
            if (Test-Path $path1) {
                $path1
            }
            elseif ($script:configDir -and (Test-Path (Join-Path $script:configDir $pageConfig.JsonPath))) {
                Join-Path $script:configDir $pageConfig.JsonPath
            }
            else {
                $path1
            }
        }
    }
    
    $sourcePath = ""
    $destPath = ""
    $logStoragePath = ""
    $logStoragePath2 = ""
    
    # ページJSONファイルが存在する場合はそこから読み込む
    if ($pageJsonPath -and (Test-Path $pageJsonPath)) {
        try {
            $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
            $sourcePath = if ($pageJson.SourcePath) { $pageJson.SourcePath } else { "" }
            # ページ3の場合はDestinationPathを読み込まない（各プロセスごとに読み込む）
            if ($script:currentPage -ne 2) {
                $destPath = if ($pageJson.DestinationPath) { $pageJson.DestinationPath } else { "" }
            }
            $logStoragePath = if ($pageJson.LogStoragePath) { $pageJson.LogStoragePath } else { "" }
            $logStoragePath2 = if ($pageJson.LogStoragePath2) { $pageJson.LogStoragePath2 } else { "" }
        }
        catch {
            Write-Log "ページJSONファイルの読み込みに失敗しました: $pageJsonPath - $($_.Exception.Message)" "ERROR"
        }
    }
    
    # 移行データファイル移動元
    # ページ3・ページ4の場合はV1抽出CSV格納元テキストボックスに設定
    if ($script:currentPage -eq 2 -or $script:currentPage -eq 3) {
        # 3ページ目・4ページ目：V1抽出CSV格納元
        # ユーザー要望により絶対パス（または既存ロジック）のまま扱う
        if ($sourcePath -and $sourcePath -ne "パス" -and $sourcePath -ne "") {
            try {
                if (-not [System.IO.Path]::IsPathRooted($sourcePath)) {
                    $sourcePath = Join-Path $PSScriptRoot $sourcePath
                }
                $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
                if ($script:v1CsvSourceTextBox) {
                    $script:v1CsvSourceTextBox.Text = $sourcePath
                }
            }
            catch {
                if ($script:v1CsvSourceTextBox) {
                    $script:v1CsvSourceTextBox.Text = "パス"
                }
            }
        }
        else {
            if ($script:v1CsvSourceTextBox) {
                $script:v1CsvSourceTextBox.Text = "パス"
            }
        }
    }
    else {
        # その他のページ：従来のsourcePathTextBox
        if ($sourcePath -and $sourcePath -ne "パス" -and $sourcePath -ne "") {
            # 相対パスの場合は絶対パスに変換
            try {
                if (-not [System.IO.Path]::IsPathRooted($sourcePath)) {
                    $sourcePath = Join-Path $PSScriptRoot $sourcePath
                }
                $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
                if ($script:sourcePathTextBox) {
                    $script:sourcePathTextBox.Text = $sourcePath
                }
            }
            catch {
                if ($script:sourcePathTextBox) {
                    $script:sourcePathTextBox.Text = "パス"
                }
            }
        }
        else {
            if ($script:sourcePathTextBox) {
                $script:sourcePathTextBox.Text = "パス"
            }
        }
    }
    
    # 移行データファイル移動先
    # ページ3の場合は各プロセス行ごとにDestinationPathを読み込むため、ここでは処理しない
    if ($script:currentPage -ne 2) {
        # その他のページ：従来のdestPathTextBox
        if ($destPath -and $destPath -ne "パス" -and $destPath -ne "") {
            # 相対パスの場合は絶対パスに変換
            try {
                if (-not [System.IO.Path]::IsPathRooted($destPath)) {
                    $destPath = Join-Path $PSScriptRoot $destPath
                }
                $destPath = [System.IO.Path]::GetFullPath($destPath)
                if ($script:destPathTextBox) {
                    $script:destPathTextBox.Text = $destPath
                }
            }
            catch {
                if ($script:destPathTextBox) {
                    $script:destPathTextBox.Text = "パス"
                }
            }
        }
        else {
            if ($script:destPathTextBox) {
                $script:destPathTextBox.Text = "パス"
            }
        }
    }
    
    # ログ格納先（1つ目）
    if ($logStoragePath -and $logStoragePath -ne "パス" -and $logStoragePath -ne "") {
        # 相対パスの場合はGlobalLogPath（なければPSScriptRoot）と結合して絶対パスに変換
        try {
            if (-not [System.IO.Path]::IsPathRooted($logStoragePath)) {
                $basePath = if ($script:globalLogPath) { $script:globalLogPath } else { $PSScriptRoot }
                $logStoragePath = Join-Path $basePath $logStoragePath
            }
            $logStoragePath = [System.IO.Path]::GetFullPath($logStoragePath)
            $script:logStoragePathTextBox.Text = $logStoragePath
        }
        catch {
            $script:logStoragePathTextBox.Text = "パス"
        }
    }
    else {
        $script:logStoragePathTextBox.Text = "パス"
    }
    
    # ログ格納先（2つ目）
    if ($null -ne $logStoragePath2 -and $logStoragePath2 -ne "" -and $logStoragePath2 -ne "パス") {
        if ($script:logStoragePath2TextBox) {
            $script:logStoragePath2TextBox.Text = $logStoragePath2
        }
    }
    else {
        if ($script:logStoragePath2TextBox) {
            $script:logStoragePath2TextBox.Text = "パス"
        }
    }
}

# ページパス保存関数
function Save-PagePaths {
    param(
        [string]$SourcePath = $null,
        [string]$DestinationPath = $null,
        [string]$LogStoragePath = $null,
        [string]$LogStoragePath2 = $null
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
    }
    else {
        $path1 = Join-Path $PSScriptRoot $pageConfig.JsonPath
        if (Test-Path $path1) {
            $path1
        }
        elseif ($script:configDir -and (Test-Path (Join-Path $script:configDir $pageConfig.JsonPath))) {
            Join-Path $script:configDir $pageConfig.JsonPath
        }
        else {
            $path1
        }
    }
    
    if (-not (Test-Path $pageJsonPath)) {
        Write-Log "ページJSONファイルが見つかりません: $pageJsonPath" "ERROR"
        return $false
    }
    
    try {
        # ページJSONファイルを読み込む
        $pageJson = Get-Content $pageJsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # SourcePath（V1抽出CSV格納元など）
        # ユーザー要望により変更なし（PSScriptRoot基準、もしくは絶対パス）
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
                }
                else {
                    $SourcePath
                }
            }
            catch {
                $SourcePath
            }
            $pageJson.SourcePath = $relativeSourcePath
        }

        # DestinationPath（V1抽出CSV格納先など）
        if ($script:currentPage -ne 2) {
            if ($DestinationPath) {
                $relativeDestPath = try {
                    # ユーザー要望により、Page 3の場合は共通基準パスを使用
                    $basePath = if ($script:currentPage -eq 2) {
                        Get-CommonBasePath
                    }
                    else {
                        [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                    }
                    
                    $targetPath = [System.IO.Path]::GetFullPath($DestinationPath).TrimEnd('\', '/')
                    
                    if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                        $relative
                    }
                    else {
                        $DestinationPath
                    }
                }
                catch {
                    $DestinationPath
                }
                $pageJson.DestinationPath = $relativeDestPath
            }
        }        
        if ($LogStoragePath) {
            $relativeLogPath = try {
                # ユーザー要望により、GlobalLogPath（ツール格納場所）からの相対パスとして保存
                $basePath = if ($script:globalLogPath) {
                    [System.IO.Path]::GetFullPath($script:globalLogPath).TrimEnd('\', '/')
                }
                else {
                    [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
                }
                $targetPath = [System.IO.Path]::GetFullPath($LogStoragePath).TrimEnd('\', '/')
                
                if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                    if ([string]::IsNullOrEmpty($relative)) {
                        $relative = "."
                    }
                    $relative
                }
                else {
                    # ユーザー要望：共通パス外の場合は保存させない（エラーメッセージを出力し、関数を中断）
                    Write-Log "ログ格納先は共通パス（$basePath）配下に配置する必要があります: $targetPath" "ERROR"
                    [void][System.Windows.Forms.MessageBox]::Show("ログ格納先は共通パス（ツール格納場所）配下に配置する必要があります。`n共通パス: $basePath`n選択されたパス: $targetPath", "設定エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return $false
                }
            }
            catch {
                Write-Log "パス変換エラー: $($_.Exception.Message)" "ERROR"
                $LogStoragePath
            }
            
            if (-not $pageJson.PSObject.Properties['LogStoragePath']) {
                $pageJson | Add-Member -MemberType NoteProperty -Name "LogStoragePath" -Value $relativeLogPath
            }
            else {
                $pageJson.LogStoragePath = $relativeLogPath
            }
        }
        else {
            if (-not $pageJson.PSObject.Properties['LogStoragePath']) {
                $pageJson | Add-Member -MemberType NoteProperty -Name "LogStoragePath" -Value ""
            }
        }
        if ($LogStoragePath2) {
            # ユーザー要望により、LogStoragePath2（ログ格納先）はフルパス（ネットワークパス対応）のまま保存
            $pageJson.LogStoragePath2 = $LogStoragePath2
            
            # LogStoragePath2プロパティが存在しない場合は追加
            if (-not $pageJson.PSObject.Properties['LogStoragePath2']) {
                $pageJson | Add-Member -MemberType NoteProperty -Name "LogStoragePath2" -Value $LogStoragePath2
            }
        }
        
        # ページJSONファイルに保存（UTF-8 BOM付き）
        $jsonContent = $pageJson | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($pageJsonPath, $jsonContent, $utf8WithBom)
        
        Write-Log "ページパスを保存しました: $pageJsonPath" "INFO"

        # メモリ上の設定も更新（動的なUI更新のため）
        if ($SourcePath) { $pageConfig | Add-Member -MemberType NoteProperty -Name "SourcePath" -Value $pageJson.SourcePath -Force }
        if ($DestinationPath) { $pageConfig | Add-Member -MemberType NoteProperty -Name "DestinationPath" -Value $pageJson.DestinationPath -Force }
        if ($LogStoragePath) { $pageConfig | Add-Member -MemberType NoteProperty -Name "LogStoragePath" -Value $pageJson.LogStoragePath -Force }
        if ($LogStoragePath2) { $pageConfig | Add-Member -MemberType NoteProperty -Name "LogStoragePath2" -Value $pageJson.LogStoragePath2 -Force }
        
        return $true
    }
    catch {
        Write-Log "ページJSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# バッチファイルパス解決関数
function Resolve-BatchPath {
    param(
        [string]$Path
    )
    
    if ([string]::IsNullOrEmpty($Path)) { return "" }
    
    # 既に絶対パスの場合はそのまま返す
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }
    
    # ユーザー指示:
    # 1. psscriptrootではなく、globallogpathを使用すること
    # 2. GlobalLogPathの親ディレクトリからの解決（パス重複回避）は不要
    
    if ($script:globalLogPath) {
        return Join-Path $script:globalLogPath $Path
    }
    
    # GlobalLogPathが未設定の場合は、パスをそのまま返す（もしくは空文字）
    return $Path
}

# プロセス実行関数
function Start-ProcessFlow {
    param([int]$ProcessIndex)
    
    # 編集モード中はファイル選択ダイアログを表示
    if ($script:editMode) {
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "バッチファイル (*.bat)|*.bat|すべてのファイル (*.*)|*.*"
        $fileDialog.Title = "バッチファイルを選択してください"
        
        # LogStoragePathを初期ディレクトリに設定
        # ユーザー要望により、GlobalLogPath（ツール格納場所）を優先使用する
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
                            # メモリ上の設定も更新（次回以降のためにキャッシュ）
                            $pageConfig.LogStoragePath = $logStoragePath
                        }
                    }
                    catch {}
                }
            }
        }

        if ($logStoragePath -and (Test-Path $logStoragePath)) {
            $fileDialog.InitialDirectory = $logStoragePath
        }
        
        # 現在のバッチファイルパスを初期値として設定（あれば）
        $currentProcesses = Get-CurrentPageProcesses
        $processConfig = $currentProcesses[$ProcessIndex]
        if ($processConfig.BatchFiles -and $processConfig.BatchFiles.Count -gt 0) {
            $currentBatch = $processConfig.BatchFiles[0]
            # Resolve-BatchPathを使用してパスを解決
            $initialPath = Resolve-BatchPath -Path $currentBatch.Path
            if (Test-Path $initialPath) {
                $fileDialog.InitialDirectory = Split-Path $initialPath
                $fileDialog.FileName = Split-Path $initialPath -Leaf
            }
        }
        
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $fileDialog.FileName
            # Save-BatchFilePath内でパス制限チェックが行われる
            if (Save-BatchFilePath -ProcessIndex $ProcessIndex -BatchFilePath $selectedFile -BatchIndex 0) {
                Write-Log "バッチファイルを設定しました: $selectedFile" "INFO" $ProcessIndex
                [System.Windows.Forms.MessageBox]::Show("バッチファイルを設定しました。`n$selectedFile", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # コントロールを更新して新しい設定を反映
                Update-ProcessControls
            }
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
            # Resolve-BatchPathを使用してパスを解決
            $batchPath = Resolve-BatchPath -Path $batch.Path
            
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
        Save-ProcessComponentExecuted -ProcessIndex $ProcessIndex -ComponentKey "ExecuteButton_Executed"
        Update-ProcessControls
    }
    else {
        Write-Log "プロセスでエラーが発生しました: $($processConfig.Name)" "ERROR" $ProcessIndex
    }
    
    $executeButton.Enabled = $true
}

# ファイル移動設定ダイアログ表示関数
function Show-FileMoveSettingsDialog {
    param(
        [int]$ProcessIndex, 
        [string]$ProcessName,
        [string]$FileSuffix = ""
    )
    
    # 現在のプロセス設定を取得
    $currentProcesses = Get-CurrentPageProcesses
    if (-not $currentProcesses -or $ProcessIndex -ge $currentProcesses.Count) {
        [System.Windows.Forms.MessageBox]::Show("プロセス情報を取得できませんでした。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $processConfig = $currentProcesses[$ProcessIndex]
    # [System.Windows.Forms.MessageBox]::Show("デバッグProcessName: $ProcessName")
    # movefiles フォルダのファイルを読み込み
    $initialText = ""
    $fileNameRaw = if ($processConfig.Name) { $processConfig.Name } elseif ($ProcessName) { $ProcessName } else { "" }   
    $fileNameRaw = $fileNameRaw.Trim()
    if (-not $fileNameRaw) {
        $fileNameRaw = "process_${script:currentPage + 1}_${ProcessIndex + 1}"
    }
    $safeFileName = [regex]::Replace($fileNameRaw, '[<>:"/\\|?*\r\n\t]', '_')
    $safeFileName += $FileSuffix
    
    $moveFilesDir = Join-Path $PSScriptRoot "config\movefiles"
    if (Test-Path $moveFilesDir) {
        $candidatePath = Join-Path $moveFilesDir ($safeFileName + ".txt")
        if (Test-Path $candidatePath) {
            try {
                $initialText = Get-Content -Path $candidatePath -Encoding UTF8 -Raw
            }
            catch {
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
        $safeFileName += $FileSuffix
        
        $moveFilesDir = Join-Path $PSScriptRoot "config\movefiles"
        if (-not (Test-Path $moveFilesDir)) {
            New-Item -ItemType Directory -Path $moveFilesDir -Force | Out-Null
        }
        $moveFilePath = Join-Path $moveFilesDir ($safeFileName + ".txt")
        Set-Content -Path $moveFilePath -Value $textBox.Text -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("ファイル移動設定を保存しました。", "保存完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    $dialogForm.Dispose()
}

# ファイルコピー共通関数（コピー元ファイル、コピー先パスを引数とする）
function Copy-FileWithLog {
    param(
        [string]$SourceFilePath,
        [string]$DestinationPath,
        [int]$ProcessIndex = -1
    )
    
    try {
        # コピー元ファイルの存在チェック
        if (-not (Test-Path $SourceFilePath)) {
            Write-Log "コピー元ファイルが見つかりません: $SourceFilePath" "ERROR" $ProcessIndex
            return $false
        }
        
        # コピー先ディレクトリが存在しない場合は作成
        if (-not (Test-Path $DestinationPath)) {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
            Write-Log "コピー先ディレクトリを作成しました: $DestinationPath" "INFO" $ProcessIndex
        }
        
        # ファイル名を取得
        $fileName = [System.IO.Path]::GetFileName($SourceFilePath)
        $destinationFilePath = Join-Path $DestinationPath $fileName
        
        # ファイルをコピー
        Copy-Item -Path $SourceFilePath -Destination $destinationFilePath -Force
        Write-Log "ファイルをコピーしました: $fileName -> $DestinationPath" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "ファイルコピーエラー: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ファイル移動実行関数（ページ3の「移動」ボタン用）
function Invoke-FileMoveOperation {
    param(
        [int]$ProcessIndex,
        [string]$ProcessName,
        [string]$V1CsvSourcePath,
        [string]$V1CsvDestinationPath,
        [string]$FileSuffix = "",
        [bool]$IsCopy = $false
    )
    
    # パラメータの検証
    if ([string]::IsNullOrWhiteSpace($ProcessName)) {
        Write-Log "プロセス名が指定されていません。" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("プロセス名が指定されていません。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($V1CsvSourcePath) -or $V1CsvSourcePath -eq "パス") {
        Write-Log "V1抽出CSV格納元が設定されていません。" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("V1抽出CSV格納元が設定されていません。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($V1CsvDestinationPath) -or $V1CsvDestinationPath -eq "パス") {
        Write-Log "V1抽出CSV格納先が設定されていません。" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("V1抽出CSV格納先が設定されていません。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # V1抽出CSV格納元の存在チェック
    if (-not (Test-Path $V1CsvSourcePath)) {
        Write-Log "V1抽出CSV格納元が存在しません: $V1CsvSourcePath" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("V1抽出CSV格納元が存在しません。`n$V1CsvSourcePath", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # movefileフォルダからファイルリストを読み込み
    $safeFileName = [regex]::Replace($ProcessName.Trim(), '[<>:"/\\|?*\r\n\t]', '_')
    $safeFileName += $FileSuffix
    $moveFilesDir = Join-Path $PSScriptRoot "config\movefiles"
    $moveFilePath = Join-Path $moveFilesDir ($safeFileName + ".txt")
    
    if (-not (Test-Path $moveFilePath)) {
        Write-Log "ファイル移動リストが見つかりません: $moveFilePath" "ERROR" $ProcessIndex
        [System.Windows.Forms.MessageBox]::Show("ファイル移動リストが見つかりません。`n先に「移動設定」でファイルリストを作成してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    Write-Log "========== ファイル移動開始 ==========" "INFO" $ProcessIndex
    Write-Log "プロセス名: $ProcessName" "INFO" $ProcessIndex
    Write-Log "V1抽出CSV格納元: $V1CsvSourcePath" "INFO" $ProcessIndex
    Write-Log "V1抽出CSV格納先: $V1CsvDestinationPath" "INFO" $ProcessIndex
    Write-Log "ファイルリスト: $moveFilePath" "INFO" $ProcessIndex
    
    # ファイルリストを1行ずつ読み込む
    $fileLines = Get-Content -Path $moveFilePath -Encoding UTF8
    $successCount = 0
    $failCount = 0
    $totalCount = 0
    
    foreach ($line in $fileLines) {
        # 空行やコメント行をスキップ
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith("#")) {
            continue
        }
        
        $totalCount++
        $fileName = $line.Trim()
        
        # V1抽出CSV格納元とファイル名を結合してコピー元のフルパスを生成
        $sourceFilePath = Join-Path $V1CsvSourcePath $fileName
        
        # ファイルコピーを実行
        $result = Copy-FileWithLog -SourceFilePath $sourceFilePath -DestinationPath $V1CsvDestinationPath -ProcessIndex $ProcessIndex
        
        if ($result) {
            $successCount++
        }
        else {
            $failCount++
        }
    }
    
    Write-Log "========== ファイル移動完了 ==========" "INFO" $ProcessIndex
    Write-Log "合計: $totalCount 件、成功: $successCount 件、失敗: $failCount 件" "INFO" $ProcessIndex
    
    # 完了メッセージを表示
    $message = "ファイル移動が完了しました。`n`n合計: $totalCount 件`n成功: $successCount 件`n失敗: $failCount 件"
    if ($failCount -gt 0) {
        [System.Windows.Forms.MessageBox]::Show($message, "移動完了（一部エラー）", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($message, "移動完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

# ログ確認関数
function Show-ProcessLog {
    param(
        [int]$ProcessIndex,
        [int]$LogIndex = 1
    )
    
    # 編集モード中はフォルダ選択ダイアログを表示
    if ($script:editMode) {
        # GlobalLogPathが設定されていない場合はエラー
        if (-not $script:globalLogPath) {
            [System.Windows.Forms.MessageBox]::Show("共通ログパス（全体設定）が設定されていません。`n先にヘッダーの「ログ格納パス設定」から設定してください。", "設定エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 初期パスの決定（GlobalLogPathを基準にする）
        $initialDirectory = $script:globalLogPath
        $currentProcesses = Get-CurrentPageProcesses
        if ($currentProcesses -and $ProcessIndex -lt $currentProcesses.Count) {
            $processConfig = $currentProcesses[$ProcessIndex]
            $propName = if ($LogIndex -eq 2) { "LogOutputDir2" } else { "LogOutputDir" }
            if ($processConfig.$propName) {
                $initialDirectory = Resolve-LogPath -SubPath $processConfig.$propName
            }
        }

        # フォルダ選択ダイアログを表示（共通関数を使用）
        $selectedPath = Show-FolderBrowser -InitialDirectory $initialDirectory -Description "ログ出力フォルダを選択してください（共通ログパス配下のみ）"

        if ($selectedPath) {
            # パスの正規化（末尾の￥削除、大文字小文字無視で比較準備）
            $normalizedGlobal = [System.IO.Path]::GetFullPath($script:globalLogPath).TrimEnd('\')
            $normalizedSelected = [System.IO.Path]::GetFullPath($selectedPath).TrimEnd('\')

            # 比較用パス（末尾に\をつける）
            $globalCompare = $normalizedGlobal
            if (-not $globalCompare.EndsWith('\')) {
                $globalCompare += '\'
            }

            # 共通パス配下かどうかチェック（完全一致 または \付きで始まるか）
            $isValid = ($normalizedSelected -eq $normalizedGlobal) -or ($normalizedSelected.StartsWith($globalCompare, [System.StringComparison]::OrdinalIgnoreCase))

            if (-not $isValid) {
                [System.Windows.Forms.MessageBox]::Show("共通ログパス配下のフォルダのみ設定可能です。`n`n共通パス: $normalizedGlobal`n選択パス: $normalizedSelected", "設定エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # 相対パスの算出
            $relativePath = $normalizedSelected.Substring($normalizedGlobal.Length).TrimStart('\')
            
            # ルートそのものが選択された場合
            if ([string]::IsNullOrEmpty($relativePath)) {
                $relativePath = "."
            }

            # JSONファイルを更新
            if (Save-ProcessLogOutputDir -ProcessIndex $ProcessIndex -LogOutputDir $relativePath -LogIndex $LogIndex) {
                Write-Log "ログ出力フォルダを設定しました: $relativePath (共通パスからの相対)" "INFO" $ProcessIndex
                [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダを設定しました。`n$relativePath", "設定完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダの保存に失敗しました。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
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
    
    # JSONファイルで設定されているLogOutputDir(2)を取得
    $processLogDir = $null
    $propName = if ($LogIndex -eq 2) { "LogOutputDir2" } else { "LogOutputDir" }
    if ($processConfig.$propName) {
        # LogOutputDir(2)が設定されている場合はそれを使用
        $processLogDir = Resolve-LogPath -SubPath $processConfig.$propName
    }
    
    # 個別設定がない場合は警告を出して終了する
    if (-not $processLogDir) {
        [System.Windows.Forms.MessageBox]::Show("このプロセスのログ出力フォルダが設定されていません。`n編集モードから「参照」ボタンをクリックして、個別のフォルダを設定してください。", "未設定", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-Log "ログ出力フォルダが未設定のため開けません: ProcessIndex=$ProcessIndex" "WARN" $ProcessIndex
        return
    }
    
    # ログ出力フォルダをエクスプローラで開く
    if (Test-Path $processLogDir) {
        # エクスプローラでフォルダを開く
        try {
            Start-Process explorer.exe -ArgumentList $processLogDir
            Write-Log "ログ出力フォルダを開きました: $processLogDir" "INFO" $ProcessIndex
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("エクスプローラを起動できませんでした。`n$processLogDir`n`n$($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "エクスプローラを起動できませんでした: $processLogDir - $($_.Exception.Message)" "ERROR" $ProcessIndex
        }
    }
    else {
        # フォルダが存在しない場合は作成してから開く
        try {
            New-Item -ItemType Directory -Path $processLogDir -Force | Out-Null
            Start-Process explorer.exe -ArgumentList $processLogDir
            Write-Log "ログ出力フォルダを作成して開きました: $processLogDir" "INFO" $ProcessIndex
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("ログ出力フォルダを開けませんでした。`n$processLogDir`n`n$($_.Exception.Message)", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "ログ出力フォルダを開けませんでした: $processLogDir - $($_.Exception.Message)" "ERROR" $ProcessIndex
        }
    }
}

# プロセスコントロールの更新

# ログ集約用バッチファイル保存関数
function Save-LogAggregationBatchFile {
    param(
        [string]$BatchFilePath
    )
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # JSONファイルパスの決定
    $jsonPath = if ($pageConfig.JsonPath) {
        if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
            $pageConfig.JsonPath
        }
        else {
            Join-Path $PSScriptRoot $pageConfig.JsonPath
        }
    }
    else {
        $null
    }
    
    if (-not $jsonPath) {
        Write-Log "JSONファイルパスが設定されていません" "ERROR"
        return $false
    }
    
    # ディレクトリ作成
    $jsonDir = Split-Path $jsonPath -Parent
    if (-not (Test-Path $jsonDir)) {
        New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null
    }
    
    # 既存のJSON読み込み or 新規作成
    $jsonContent = if (Test-Path $jsonPath) {
        try {
            Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        }
        catch {
            New-Object PSObject
        }
    }
    else {
        New-Object PSObject
    }
    
    # 相対パス変換
    $relativeBatchPath = try {
        $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
        $targetPath = [System.IO.Path]::GetFullPath($BatchFilePath).TrimEnd('\', '/')
        
        if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
            if ([string]::IsNullOrEmpty($relative)) {
                $relative = Split-Path $targetPath -Leaf
            }
            $relative
        }
        else {
            $BatchFilePath
        }
    }
    catch {
        $BatchFilePath
    }
    
    # LogAggregationBatchFileオブジェクト作成
    $batchObj = New-Object PSObject
    $batchObj | Add-Member -MemberType NoteProperty -Name "Name" -Value (Split-Path $BatchFilePath -Leaf)
    $batchObj | Add-Member -MemberType NoteProperty -Name "Path" -Value $relativeBatchPath
    
    # プロパティ追加/更新
    if (-not $jsonContent.PSObject.Properties['LogAggregationBatchFile']) {
        $jsonContent | Add-Member -MemberType NoteProperty -Name "LogAggregationBatchFile" -Value $batchObj
    }
    else {
        $jsonContent.LogAggregationBatchFile = $batchObj
    }
    
    try {
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        Write-Log "ログ集約用バッチファイルを設定しました: $BatchFilePath" "INFO"
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# プロセス有効/無効保存関数
function Save-ProcessEnabled {
    param(
        [int]$ProcessIndex,
        [bool]$Enabled
    )
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "このページはJSONファイルを使用していません" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    }
    else {
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
        
        $procObj = $jsonContent.Processes[$ProcessIndex]
        
        # Enabledプロパティを強制的に追加/更新
        $procObj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $Enabled -Force
        
        # プロセスが有効化された場合、個別ボタンの実行済みフラグもリセットする
        if ($Enabled) {
            $executedProperties = $procObj.PSObject.Properties | Where-Object { $_.Name -like "*_Executed" }
            foreach ($prop in $executedProperties) {
                $procObj | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $false -Force
            }
        }
        
        # JSONファイルに保存（UTF-8 BOM付き）
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        
        # メモリ上の設定も更新（重要：UI再描画時のため）
        if ($script:pages[$script:currentPage].Processes -and $ProcessIndex -lt $script:pages[$script:currentPage].Processes.Count) {
            $memProc = $script:pages[$script:currentPage].Processes[$ProcessIndex]
            $memProc | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $Enabled -Force
            if ($Enabled) {
                $mExecutedProperties = $memProc.PSObject.Properties | Where-Object { $_.Name -like "*_Executed" }
                foreach ($mProp in $mExecutedProperties) {
                    $memProc | Add-Member -MemberType NoteProperty -Name $mProp.Name -Value $false -Force
                }
            }
        }
        
        Write-Log "プロセス有効状態を保存しました: $Enabled" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "JSONファイルの保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# プロセス内コンポーネントの実行済み状態保存関数
function Save-ProcessComponentExecuted {
    param(
        [int]$ProcessIndex,
        [string]$ComponentKey,
        [bool]$Executed = $true
    )
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) { return $false }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) { $pageConfig.JsonPath } else { Join-Path $PSScriptRoot $pageConfig.JsonPath }
    if (-not (Test-Path $jsonPath)) { return $false }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not $jsonContent.Processes -or $ProcessIndex -ge $jsonContent.Processes.Count) { return $false }
        
        $procObj = $jsonContent.Processes[$ProcessIndex]
        
        # 実行済みフラグを更新
        $procObj | Add-Member -MemberType NoteProperty -Name $ComponentKey -Value $Executed -Force
        
        # JSONファイルに保存
        $jsonContentStr = $jsonContent | ConvertTo-Json -Depth 10
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($jsonPath, $jsonContentStr, $utf8WithBom)
        
        # メモリ上の設定も更新
        if ($script:pages[$script:currentPage].Processes -and $ProcessIndex -lt $script:pages[$script:currentPage].Processes.Count) {
            $memProc = $script:pages[$script:currentPage].Processes[$ProcessIndex]
            $memProc | Add-Member -MemberType NoteProperty -Name $ComponentKey -Value $Executed -Force
        }
        
        Write-Log "コンポーネント実行状態を保存しました: $ComponentKey = $Executed" "INFO" $ProcessIndex
        return $true
    }
    catch {
        Write-Log "コンポーネント実行状態の保存に失敗しました: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

