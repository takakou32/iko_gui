# PowerShell�X�N���v�g - GUI�A�v���P�[�V����
# �G���R�[�f�B���O: Shift-JIS

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �ݒ�t�@�C���̓ǂݍ���
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json
} else {
    Write-Host "�ݒ�t�@�C����������܂���: $configPath"
    exit 1
}

# ���O�t�@�C���̃p�X
$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# �O���[�o���ϐ�
$script:currentPage = 0
$script:processesPerPage = 8
$script:processControls = @()
$script:processLogs = @{}
$script:pages = @()
$script:pageProcessCache = @()
$script:editMode = $false

# �y�[�W�ݒ�̓ǂݍ���
if ($config.Pages) {
    $script:pages = $config.Pages
} else {
    # ����݊����̂��߁A���`���̐ݒ���T�|�[�g
    if ($config.Processes) {
        $script:pages = @(@{
            Title = if ($config.Title) { $config.Title } else { "" }
            JsonPath = $null
            Processes = $config.Processes
        })
    } else {
        Write-Host "�ݒ�t�@�C���̌`��������������܂���"
        exit 1
    }
}

# ���݂̃y�[�W�̃v���Z�X�ꗗ���擾
function Get-CurrentPageProcesses {
    if ($script:currentPage -ge $script:pages.Count) {
        return @()
    }
    
    $pageConfig = $script:pages[$script:currentPage]
    
    # JsonPath���w�肳��Ă���ꍇ�́A����JSON�t�@�C����ǂݍ���
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
                    Write-Log "JSON�t�@�C����Processes���܂܂�Ă��܂���: $jsonPath" "WARN"
                    return @()
                }
            } catch {
                Write-Log "JSON�t�@�C���̓ǂݍ��݂Ɏ��s���܂���: $jsonPath - $($_.Exception.Message)" "ERROR"
                return @()
            }
        } else {
            Write-Log "JSON�t�@�C����������܂���: $jsonPath" "ERROR"
            return @()
        }
    }
    
    # JsonPath���w�肳��Ă��Ȃ��ꍇ�́A����Processes���g�p�i����݊����j
    if ($pageConfig.Processes) {
        return $pageConfig.Processes
    }
    
    return @()
}

# ���O�o�͊֐�
function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [int]$ProcessIndex = -1, [string]$LogDir = $null)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # �v���Z�X�ŗL�̃��O�t�@�C��
    if ($ProcessIndex -ge 0) {
        # LogDir���w�肳��Ă��Ȃ��ꍇ�AProcessIndex����擾
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
        
        # ���O�f�B���N�g�������݂��Ȃ��ꍇ�͍쐬
        if (-not (Test-Path $LogDir)) {
            New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
        }
        
        $processLogFile = Join-Path $LogDir "process_${script:currentPage}_${ProcessIndex}.log"
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::AppendAllText($processLogFile, $logMessage + "`r`n", $utf8NoBom)
        $script:processLogs["${script:currentPage}_${ProcessIndex}"] = $processLogFile
    }
    
    # GUI�̃��O�\���G���A�ɒǉ�
    $script:logTextBox.AppendText("$logMessage`r`n")
    $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length
    $script:logTextBox.ScrollToCaret()
    
    Write-Host $logMessage
}

# Bat�t�@�C�����s�֐�
function Invoke-BatchFile {
    param(
        [string]$BatchPath,
        [string]$DisplayName,
        [int]$ProcessIndex
    )
    
    # �p�X�̐��K������
    if ([string]::IsNullOrWhiteSpace($BatchPath)) {
        Write-Log "�o�b�`�t�@�C���p�X����ł�" "ERROR" $ProcessIndex
        return $false
    }
    
    # �擪�E�����̋󔒂��폜
    $BatchPath = $BatchPath.Trim()
    
    # �p�X�𐳋K���i���΃p�X�̉����A��؂蕶���̓���Ȃǁj
    try {
        # ���΃p�X�̏ꍇ��$PSScriptRoot����ɉ���
        if (-not [System.IO.Path]::IsPathRooted($BatchPath)) {
            $BatchPath = Join-Path $PSScriptRoot $BatchPath
        }
        # �p�X�𐳋K���i..��.�������A��؂蕶���𓝈�j
        $BatchPath = [System.IO.Path]::GetFullPath($BatchPath)
    } catch {
        Write-Log "�o�b�`�t�@�C���p�X�̐��K���Ɏ��s���܂���: $BatchPath - $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
    
    if (-not (Test-Path $BatchPath)) {
        Write-Log "�o�b�`�t�@�C����������܂���: $BatchPath" "ERROR" $ProcessIndex
        return $false
    }
    
    # ���O�o�̓f�B���N�g���̌���
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
    
    Write-Log "�o�b�`�t�@�C�������s��: $DisplayName ($BatchPath)" "INFO" $ProcessIndex
    
    try {
        $stdoutFile = Join-Path $processLogDir "process_${script:currentPage}_${ProcessIndex}_stdout.log"
        $stderrFile = Join-Path $processLogDir "process_${script:currentPage}_${ProcessIndex}_stderr.log"
        
        $process = Start-Process -FilePath $BatchPath -WorkingDirectory (Split-Path $BatchPath) -Wait -NoNewWindow -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
        
        if ($process.ExitCode -eq 0) {
            Write-Log "�o�b�`�t�@�C���̎��s���������܂���: $DisplayName (�I���R�[�h: $($process.ExitCode))" "INFO" $ProcessIndex
            return $true
        } else {
            Write-Log "�o�b�`�t�@�C���̎��s�ŃG���[���������܂���: $DisplayName (�I���R�[�h: $($process.ExitCode))" "ERROR" $ProcessIndex
            return $false
        }
    } catch {
        Write-Log "�o�b�`�t�@�C���̎��s���ɗ�O���������܂���: $DisplayName - $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# CSV�t�@�C���ړ��֐�
function Move-CsvFiles {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$ProcessIndex
    )
    
    if (-not (Test-Path $SourcePath)) {
        Write-Log "�\�[�X�p�X��������܂���: $SourcePath" "ERROR" $ProcessIndex
        return $false
    }
    
    if (-not (Test-Path $DestinationPath)) {
        Write-Log "�ړ���f�B���N�g�����쐬���܂�: $DestinationPath" "INFO" $ProcessIndex
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
    
    try {
        $csvFiles = Get-ChildItem -Path $SourcePath -Filter "*.csv" -File
        
        if ($csvFiles.Count -eq 0) {
            Write-Log "CSV�t�@�C����������܂���: $SourcePath" "WARN" $ProcessIndex
            return $false
        }
        
        foreach ($file in $csvFiles) {
            $destFile = Join-Path $DestinationPath $file.Name
            Move-Item -Path $file.FullName -Destination $destFile -Force
            Write-Log "CSV�t�@�C�����ړ����܂���: $($file.Name) -> $DestinationPath" "INFO" $ProcessIndex
        }
        
        Write-Log "CSV�t�@�C���̈ړ����������܂��� (�ړ���: $($csvFiles.Count))" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "CSV�t�@�C���̈ړ����ɃG���[���������܂���: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# �o�b�`�t�@�C���p�X�ۑ��֐�
function Save-BatchFilePath {
    param([int]$ProcessIndex, [string]$BatchFilePath, [int]$BatchIndex = 0)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "���̃y�[�W��JSON�t�@�C�����g�p���Ă��܂���" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    } else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $jsonPath)) {
        Write-Log "JSON�t�@�C����������܂���: $jsonPath" "ERROR" $ProcessIndex
        return $false
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not $jsonContent.Processes -or $ProcessIndex -ge $jsonContent.Processes.Count) {
            Write-Log "�v���Z�X�C���f�b�N�X���͈͊O�ł�" "ERROR" $ProcessIndex
            return $false
        }
        
        $process = $jsonContent.Processes[$ProcessIndex]
        if (-not $process.BatchFiles) {
            $process.BatchFiles = @()
        }
        
        if ($BatchIndex -ge $process.BatchFiles.Count) {
            # �V�����o�b�`�t�@�C���G���g����ǉ�
            $process.BatchFiles += @{
                Name = "�o�b�`�t�@�C��"
                Path = $BatchFilePath
            }
        } else {
            # �����̃o�b�`�t�@�C���G���g�����X�V
            $process.BatchFiles[$BatchIndex].Path = $BatchFilePath
        }
        
        # ���΃p�X�ɕϊ��i�\�ȏꍇ�j
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
        
        # JSON�t�@�C���ɕۑ�
        $jsonContent | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
        Write-Log "�o�b�`�t�@�C���p�X��ۑ����܂���: $relativePath" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "JSON�t�@�C���̕ۑ��Ɏ��s���܂���: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# ���O�o�̓t�H���_�p�X�ۑ��֐�
function Save-ProcessLogOutputDir {
    param([int]$ProcessIndex, [string]$LogOutputDir)
    
    $pageConfig = $script:pages[$script:currentPage]
    if (-not $pageConfig.JsonPath) {
        Write-Log "���̃y�[�W��JSON�t�@�C�����g�p���Ă��܂���" "WARN" $ProcessIndex
        return $false
    }
    
    $jsonPath = if ([System.IO.Path]::IsPathRooted($pageConfig.JsonPath)) {
        $pageConfig.JsonPath
    } else {
        Join-Path $PSScriptRoot $pageConfig.JsonPath
    }
    
    if (-not (Test-Path $jsonPath)) {
        Write-Log "JSON�t�@�C����������܂���: $jsonPath" "ERROR" $ProcessIndex
        return $false
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not $jsonContent.Processes -or $ProcessIndex -ge $jsonContent.Processes.Count) {
            Write-Log "�v���Z�X�C���f�b�N�X���͈͊O�ł�" "ERROR" $ProcessIndex
            return $false
        }
        
        $process = $jsonContent.Processes[$ProcessIndex]
        $process.LogOutputDir = $LogOutputDir
        
        # JSON�t�@�C���ɕۑ�
        $jsonContent | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
        Write-Log "���O�o�̓t�H���_�p�X��ۑ����܂���: $LogOutputDir" "INFO" $ProcessIndex
        return $true
    } catch {
        Write-Log "JSON�t�@�C���̕ۑ��Ɏ��s���܂���: $($_.Exception.Message)" "ERROR" $ProcessIndex
        return $false
    }
}

# config.json�ۑ��֐�
function Save-ConfigFile {
    try {
        $configPath = Join-Path $PSScriptRoot "config.json"
        # $script:pages�̕ύX��$config�ɔ��f
        $config.Pages = $script:pages
        $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
        return $true
    } catch {
        Write-Log "config.json�̕ۑ��Ɏ��s���܂���: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# �y�[�W�p�X�ۑ��֐�
function Save-PagePath {
    param(
        [string]$PathType,
        [string]$Path
    )
    
    if ($script:currentPage -ge $script:pages.Count) {
        Write-Log "�y�[�W�C���f�b�N�X���͈͊O�ł�" "ERROR"
        return $false
    }
    
    try {
        $pageConfig = $script:pages[$script:currentPage]
        
        # ���΃p�X�ɕϊ��i�\�ȏꍇ�j
        $relativePath = try {
            $basePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/')
            $targetPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
            
            if ($targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relative = $targetPath.Substring($basePath.Length).TrimStart('\', '/')
                if ([string]::IsNullOrEmpty($relative)) {
                    $relative = Split-Path $targetPath -Leaf
                }
                $relative
            } else {
                $Path
            }
        } catch {
            $Path
        }
        
        # �p�X��ݒ�
        switch ($PathType) {
            "MigrationSource" {
                $pageConfig.MigrationSourcePath = $relativePath
            }
            "MigrationDest" {
                $pageConfig.MigrationDestPath = $relativePath
            }
            "LogStorage" {
                $pageConfig.LogStoragePath = $relativePath
            }
            default {
                Write-Log "�s���ȃp�X�^�C�v�ł�: $PathType" "ERROR"
                return $false
            }
        }
        
        # config.json�ɕۑ�
        if (Save-ConfigFile) {
            Write-Log "$PathType �p�X��ۑ����܂���: $relativePath" "INFO"
            return $true
        } else {
            return $false
        }
    } catch {
        Write-Log "�y�[�W�p�X�̕ۑ��Ɏ��s���܂���: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# �v���Z�X���s�֐�
function Start-ProcessFlow {
    param([int]$ProcessIndex)
    
    # �ҏW���[�h���̓t�@�C���I���_�C�A���O��\��
    if ($script:editMode) {
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "�o�b�`�t�@�C�� (*.bat)|*.bat|���ׂẴt�@�C�� (*.*)|*.*"
        $fileDialog.Title = "�o�b�`�t�@�C����I�����Ă�������"
        
        # ���݂̃o�b�`�t�@�C���p�X�������l�Ƃ��Đݒ�
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
            Write-Log "�o�b�`�t�@�C����ݒ肵�܂���: $selectedFile" "INFO" $ProcessIndex
            [System.Windows.Forms.MessageBox]::Show("�o�b�`�t�@�C����ݒ肵�܂����B`n$selectedFile", "�ݒ芮��", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # �R���g���[�����X�V���ĐV�����ݒ�𔽉f
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
    
    Write-Log "�v���Z�X���J�n���܂�: $($processConfig.Name)" "INFO" $ProcessIndex
    
    $allSuccess = $true
    
    # �o�b�`�t�@�C���̎��s
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
            
            # ���s�Ԋu�i�ݒ肳��Ă���ꍇ�j
            if ($processConfig.ExecutionDelay -and $processConfig.ExecutionDelay -gt 0) {
                Start-Sleep -Seconds $processConfig.ExecutionDelay
            }
        }
    }
    
    # CSV�t�@�C���̈ړ�
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
        Write-Log "�v���Z�X������Ɋ������܂���: $($processConfig.Name)" "INFO" $ProcessIndex
    } else {
        Write-Log "�v���Z�X�ŃG���[���������܂���: $($processConfig.Name)" "ERROR" $ProcessIndex
    }
    
    $executeButton.Enabled = $true
}

# ���O�m�F�֐�
function Show-ProcessLog {
    param([int]$ProcessIndex)
    
    # �ҏW���[�h���̓t�H���_�I���_�C�A���O��\��
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "���O�o�̓t�H���_��I�����Ă�������"
        $folderDialog.ShowNewFolderButton = $true
        
        # ���݂̃��O�t�H���_�������l�Ƃ��Đݒ�
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
            # ���΃p�X�ɕϊ��i�\�ȏꍇ�j
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
            
            # JSON�t�@�C�����X�V
            Save-ProcessLogOutputDir -ProcessIndex $ProcessIndex -LogOutputDir $relativePath
            Write-Log "���O�o�̓t�H���_��ݒ肵�܂���: $relativePath" "INFO" $ProcessIndex
            [System.Windows.Forms.MessageBox]::Show("���O�o�̓t�H���_��ݒ肵�܂����B`n$relativePath", "�ݒ芮��", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        $folderDialog.Dispose()
        return
    }
    
    # �ʏ탂�[�h�ł̓��O�t�@�C�����J��
    $currentProcesses = Get-CurrentPageProcesses
    $processConfig = $currentProcesses[$ProcessIndex]
    $processLogDir = $logDir
    if ($processConfig.LogOutputDir) {
        $processLogDir = if ([System.IO.Path]::IsPathRooted($processConfig.LogOutputDir)) {
            $processConfig.LogOutputDir
        } else {
            Join-Path $PSScriptRoot $processConfig.LogOutputDir
        }
    }
    
    $logKey = "${script:currentPage}_${ProcessIndex}"
    if ($script:processLogs.ContainsKey($logKey) -and (Test-Path $script:processLogs[$logKey])) {
        Start-Process notepad.exe -ArgumentList $script:processLogs[$logKey]
    } else {
        $processLogFile = Join-Path $processLogDir "process_${script:currentPage}_${ProcessIndex}.log"
        if (Test-Path $processLogFile) {
            Start-Process notepad.exe -ArgumentList $processLogFile
        } else {
            [System.Windows.Forms.MessageBox]::Show("���O�t�@�C����������܂���B", "�G���[", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
}

# �v���Z�X�R���g���[���̍X�V
function Update-ProcessControls {
    # �y�[�W�J�ڎ���JSON�t�@�C����ǂݍ���
    $currentProcesses = Get-CurrentPageProcesses
    $totalPages = $script:pages.Count
    
    Write-Log "�y�[�W $($script:currentPage + 1) �̃v���Z�X��ǂݍ��݂܂��� (�v���Z�X��: $($currentProcesses.Count))" "INFO"
    
    # �����̃R���g���[�����N���A
    foreach ($ctrlGroup in $script:processControls) {
        if ($ctrlGroup) {
            $script:processPanel.Controls.Remove($ctrlGroup.NameTextBox)
            $script:processPanel.Controls.Remove($ctrlGroup.ExecuteButton)
            $script:processPanel.Controls.Remove($ctrlGroup.LogButton)
        }
    }
    $script:processControls = @()
    
    # �V�����R���g���[�����쐬
    for ($i = 0; $i -lt $script:processesPerPage; $i++) {
        if ($i -lt $currentProcesses.Count) {
            $processConfig = $currentProcesses[$i]
            $row = [Math]::Floor($i / 2)
            $col = $i % 2
            
            # �R���g���[���̈ʒu�v�Z
            $x = [int](10 + $col * 390)
            $y = [int](10 + $row * 60)
            
            # �e�L�X�g�{�b�N�X�i�^�X�N���\���p�j
            $nameTextBox = New-Object System.Windows.Forms.TextBox
            $nameTextBox.Location = New-Object System.Drawing.Point($x, $y)
            $nameTextBox.Size = New-Object System.Drawing.Size(200, 40)
            $nameTextBox.Text = if ($processConfig.Name) { $processConfig.Name } else { "" }
            $nameTextBox.ReadOnly = $true
            $nameTextBox.BackColor = [System.Drawing.Color]::White
            $nameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $nameTextBox.Font = New-Object System.Drawing.Font("���C���I", 9)
            $nameTextBox.Multiline = $true
            $nameTextBox.Height = 40
            $script:processPanel.Controls.Add($nameTextBox)
            
            # ���s�{�^���i�I�����W�j
            $executeButton = New-Object System.Windows.Forms.Button
            $executeX = [int]($x + 210)
            $executeButton.Location = New-Object System.Drawing.Point($executeX, $y)
            $executeButton.Size = New-Object System.Drawing.Size(80, 40)
            if ($script:editMode) {
                $executeButton.Text = "�Q��"
            } else {
                $executeButton.Text = if ($processConfig.ExecuteButtonText) { $processConfig.ExecuteButtonText } else { "���s" }
            }
            $executeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
            $executeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $executeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
            $executeButton.FlatAppearance.BorderSize = 1
            $executeButton.Font = New-Object System.Drawing.Font("���C���I", 9)
            $processIdx = $i
            $executeButton.Add_Click({
                Start-ProcessFlow -ProcessIndex $processIdx
            })
            $script:processPanel.Controls.Add($executeButton)
            
            # ���O�m�F�{�^���i�΁j
            $logButton = New-Object System.Windows.Forms.Button
            $logX = [int]($x + 300)
            $logButton.Location = New-Object System.Drawing.Point($logX, $y)
            $logButton.Size = New-Object System.Drawing.Size(80, 40)
            if ($script:editMode) {
                $logButton.Text = "�Q��"
            } else {
                $logButton.Text = if ($processConfig.LogButtonText) { $processConfig.LogButtonText } else { "���O�m�F" }
            }
            $logButton.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 200)
            $logButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $logButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
            $logButton.FlatAppearance.BorderSize = 1
            $logButton.Font = New-Object System.Drawing.Font("���C���I", 9)
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
    
    # �y�[�W���̍X�V
    $script:pageLabel.Text = "�y�[�W $($script:currentPage + 1) / $totalPages"
    
    # �^�C�g���̍X�V
    $pageTitle = if ($script:pages[$script:currentPage].Title) { $script:pages[$script:currentPage].Title } else { if ($config.Title) { $config.Title } else { "1.V1 �ڍs�c�[���K�p" } }
    $script:titleLabel.Text = $pageTitle
    
    # �y�[�W���Ƃ̃p�X����ǂݍ���ŕ\��
    $pageConfig = $script:pages[$script:currentPage]
    
    # �ڍs�f�[�^�t�@�C���ړ���
    if ($pageConfig.MigrationSourcePath) {
        $sourcePath = if ([System.IO.Path]::IsPathRooted($pageConfig.MigrationSourcePath)) {
            $pageConfig.MigrationSourcePath
        } else {
            Join-Path $PSScriptRoot $pageConfig.MigrationSourcePath
        }
        if (Test-Path $sourcePath) {
            $script:sourceTextBox.Text = $sourcePath
        } else {
            $script:sourceTextBox.Text = $pageConfig.MigrationSourcePath
        }
    } else {
        $script:sourceTextBox.Text = ""
    }
    
    # �ڍs�f�[�^�t�@�C���ړ���
    if ($pageConfig.MigrationDestPath) {
        $destPath = if ([System.IO.Path]::IsPathRooted($pageConfig.MigrationDestPath)) {
            $pageConfig.MigrationDestPath
        } else {
            Join-Path $PSScriptRoot $pageConfig.MigrationDestPath
        }
        if (Test-Path $destPath) {
            $script:destTextBox.Text = $destPath
        } else {
            $script:destTextBox.Text = $pageConfig.MigrationDestPath
        }
    } else {
        $script:destTextBox.Text = ""
    }
    
    # ���O�i�[��
    if ($pageConfig.LogStoragePath) {
        $logPath = if ([System.IO.Path]::IsPathRooted($pageConfig.LogStoragePath)) {
            $pageConfig.LogStoragePath
        } else {
            Join-Path $PSScriptRoot $pageConfig.LogStoragePath
        }
        if (Test-Path $logPath) {
            $script:logStorageTextBox.Text = $logPath
        } else {
            $script:logStorageTextBox.Text = $pageConfig.LogStoragePath
        }
    } else {
        $script:logStorageTextBox.Text = ""
    }
}

# GUI�t�H�[���̍쐬
$form = New-Object System.Windows.Forms.Form
$form.Text = "�v���Z�X���sGUI"
$form.Size = New-Object System.Drawing.Size(800, 1000)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

# �w�b�_�[�����i���F�w�i�j
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(800, 50)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
$form.Controls.Add($headerPanel)

# �^�C�g�����x��
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
$titleLabel.Text = if ($script:pages.Count -gt 0 -and $script:pages[0].Title) { $script:pages[0].Title } else { if ($config.Title) { $config.Title } else { "1.V1 �ڍs�c�[���K�p" } }
$titleLabel.Font = New-Object System.Drawing.Font("���C���I", 12, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($titleLabel)
$script:titleLabel = $titleLabel

    # �����{�^��
$leftArrowButton = New-Object System.Windows.Forms.Button
$leftArrowButton.Location = New-Object System.Drawing.Point(690, 10)
$leftArrowButton.Size = New-Object System.Drawing.Size(40, 30)
$leftArrowButton.Text = "<"
$leftArrowButton.BackColor = [System.Drawing.Color]::Black
$leftArrowButton.ForeColor = [System.Drawing.Color]::White
$leftArrowButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$leftArrowButton.Font = New-Object System.Drawing.Font("���C���I", 12, [System.Drawing.FontStyle]::Bold)
$leftArrowButton.Add_Click({
    if ($script:currentPage -gt 0) {
        $script:currentPage--
        Update-ProcessControls
    }
})
$headerPanel.Controls.Add($leftArrowButton)

# �E���{�^��
$rightArrowButton = New-Object System.Windows.Forms.Button
$rightArrowButton.Location = New-Object System.Drawing.Point(740, 10)
$rightArrowButton.Size = New-Object System.Drawing.Size(40, 30)
$rightArrowButton.Text = ">"
$rightArrowButton.BackColor = [System.Drawing.Color]::Black
$rightArrowButton.ForeColor = [System.Drawing.Color]::White
$rightArrowButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rightArrowButton.Font = New-Object System.Drawing.Font("���C���I", 12, [System.Drawing.FontStyle]::Bold)
$rightArrowButton.Add_Click({
    if ($script:currentPage -lt ($script:pages.Count - 1)) {
        $script:currentPage++
        Update-ProcessControls
    }
})
$headerPanel.Controls.Add($rightArrowButton)

# �y�[�W���x��
$pageLabel = New-Object System.Windows.Forms.Label
$pageLabel.Location = New-Object System.Drawing.Point(420, 10)
$pageLabel.Size = New-Object System.Drawing.Size(150, 30)
$pageLabel.Text = "�y�[�W 1 / $($script:pages.Count)"
$pageLabel.Font = New-Object System.Drawing.Font("���C���I", 10)
$pageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$headerPanel.Controls.Add($pageLabel)
$script:pageLabel = $pageLabel

# �ҏW���[�h�؂�ւ��{�^��
$editModeButton = New-Object System.Windows.Forms.Button
$editModeButton.Location = New-Object System.Drawing.Point(580, 10)
$editModeButton.Size = New-Object System.Drawing.Size(100, 30)
$editModeButton.Text = "�ҏW���[�h OFF"
$editModeButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$editModeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$editModeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$editModeButton.FlatAppearance.BorderSize = 1
$editModeButton.Font = New-Object System.Drawing.Font("���C���I", 9)
$editModeButton.Add_Click({
    $script:editMode = -not $script:editMode
    if ($script:editMode) {
        $editModeButton.Text = "�ҏW���[�h ON"
        $editModeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
        Write-Log "�ҏW���[�h��L���ɂ��܂���" "INFO"
    } else {
        $editModeButton.Text = "�ҏW���[�h OFF"
        $editModeButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        Write-Log "�ҏW���[�h�𖳌��ɂ��܂���" "INFO"
    }
    # �{�^���̃e�L�X�g���X�V
    Update-ProcessControls
})
$headerPanel.Controls.Add($editModeButton)
$script:editModeButton = $editModeButton

# �v���Z�X����G���A�i���F/�x�[�W���w�i�j
$processPanel = New-Object System.Windows.Forms.Panel
$processPanel.Location = New-Object System.Drawing.Point(0, 50)
$processPanel.Size = New-Object System.Drawing.Size(800, 280)
$processPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
$form.Controls.Add($processPanel)
$script:processPanel = $processPanel

# ���O�\���G���A�͌�ō쐬�i���O�i�[��̉��ɔz�u�j

# �ڍs�f�[�^�t�@�C���ړ���
$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Location = New-Object System.Drawing.Point(10, 570)
$sourceLabel.Size = New-Object System.Drawing.Size(200, 20)
$sourceLabel.Text = "�ڍs�f�[�^�t�@�C���ړ���"
$sourceLabel.Font = New-Object System.Drawing.Font("���C���I", 9)
$form.Controls.Add($sourceLabel)

$sourceTextBox = New-Object System.Windows.Forms.TextBox
$sourceTextBox.Location = New-Object System.Drawing.Point(10, 590)
$sourceTextBox.Size = New-Object System.Drawing.Size(500, 25)
$sourceTextBox.Font = New-Object System.Drawing.Font("���C���I", 9)
$sourceTextBox.PlaceholderText = "�p�X"
$sourceTextBox.ReadOnly = $true
$sourceTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "�ڍs�f�[�^�t�@�C���ړ����t�H���_��I�����Ă�������"
        $folderDialog.ShowNewFolderButton = $true
        
        # ���݂̃p�X�������l�Ƃ��Đݒ�
        if ($sourceTextBox.Text -and (Test-Path $sourceTextBox.Text)) {
            $folderDialog.SelectedPath = $sourceTextBox.Text
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $sourceTextBox.Text = $folderDialog.SelectedPath
            Save-PagePath -PathType "MigrationSource" -Path $folderDialog.SelectedPath
        }
        $folderDialog.Dispose()
    }
})
$form.Controls.Add($sourceTextBox)
$script:sourceTextBox = $sourceTextBox

# �ڍs�f�[�^�t�@�C���ړ���
$destLabel = New-Object System.Windows.Forms.Label
$destLabel.Location = New-Object System.Drawing.Point(10, 625)
$destLabel.Size = New-Object System.Drawing.Size(200, 20)
$destLabel.Text = "�ڍs�f�[�^�t�@�C���ړ���"
$destLabel.Font = New-Object System.Drawing.Font("���C���I", 9)
$form.Controls.Add($destLabel)

$destTextBox = New-Object System.Windows.Forms.TextBox
$destTextBox.Location = New-Object System.Drawing.Point(10, 645)
$destTextBox.Size = New-Object System.Drawing.Size(500, 25)
$destTextBox.Font = New-Object System.Drawing.Font("���C���I", 9)
$destTextBox.PlaceholderText = "�p�X"
$destTextBox.ReadOnly = $true
$destTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "�ڍs�f�[�^�t�@�C���ړ���t�H���_��I�����Ă�������"
        $folderDialog.ShowNewFolderButton = $true
        
        # ���݂̃p�X�������l�Ƃ��Đݒ�
        if ($destTextBox.Text -and (Test-Path $destTextBox.Text)) {
            $folderDialog.SelectedPath = $destTextBox.Text
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destTextBox.Text = $folderDialog.SelectedPath
            Save-PagePath -PathType "MigrationDest" -Path $folderDialog.SelectedPath
        }
        $folderDialog.Dispose()
    }
})
$form.Controls.Add($destTextBox)
$script:destTextBox = $destTextBox

$fileMoveButton = New-Object System.Windows.Forms.Button
$fileMoveButton.Location = New-Object System.Drawing.Point(520, 645)
$fileMoveButton.Size = New-Object System.Drawing.Size(90, 25)
$fileMoveButton.Text = "�t�@�C���ړ�"
$fileMoveButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
$fileMoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$fileMoveButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$fileMoveButton.FlatAppearance.BorderSize = 1
$fileMoveButton.Font = New-Object System.Drawing.Font("���C���I", 9)
$fileMoveButton.Add_Click({
    # �t�@�C���ړ������������Ɏ���
    Write-Log "�t�@�C���ړ����������s���܂�" "INFO"
})
$form.Controls.Add($fileMoveButton)
$script:fileMoveButton = $fileMoveButton

# ���O�i�[��
$logStorageLabel = New-Object System.Windows.Forms.Label
$logStorageLabel.Location = New-Object System.Drawing.Point(10, 680)
$logStorageLabel.Size = New-Object System.Drawing.Size(200, 20)
$logStorageLabel.Text = "���O�i�[��"
$logStorageLabel.Font = New-Object System.Drawing.Font("���C���I", 9)
$form.Controls.Add($logStorageLabel)

$logStorageTextBox = New-Object System.Windows.Forms.TextBox
$logStorageTextBox.Location = New-Object System.Drawing.Point(10, 700)
$logStorageTextBox.Size = New-Object System.Drawing.Size(500, 25)
$logStorageTextBox.Font = New-Object System.Drawing.Font("���C���I", 9)
$logStorageTextBox.PlaceholderText = "�p�X"
$logStorageTextBox.ReadOnly = $true
$logStorageTextBox.Add_Click({
    if ($script:editMode) {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "���O�i�[��t�H���_��I�����Ă�������"
        $folderDialog.ShowNewFolderButton = $true
        
        # ���݂̃p�X�������l�Ƃ��Đݒ�
        if ($logStorageTextBox.Text -and (Test-Path $logStorageTextBox.Text)) {
            $folderDialog.SelectedPath = $logStorageTextBox.Text
        } elseif (Test-Path $logDir) {
            $folderDialog.SelectedPath = $logDir
        } elseif (Test-Path $PSScriptRoot) {
            $folderDialog.SelectedPath = $PSScriptRoot
        }
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $logStorageTextBox.Text = $folderDialog.SelectedPath
            Save-PagePath -PathType "LogStorage" -Path $folderDialog.SelectedPath
        }
        $folderDialog.Dispose()
    }
})
$form.Controls.Add($logStorageTextBox)
$script:logStorageTextBox = $logStorageTextBox

$logStoreButton = New-Object System.Windows.Forms.Button
$logStoreButton.Location = New-Object System.Drawing.Point(520, 700)
$logStoreButton.Size = New-Object System.Drawing.Size(90, 25)
$logStoreButton.Text = "���O�i�["
$logStoreButton.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)
$logStoreButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$logStoreButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
$logStoreButton.FlatAppearance.BorderSize = 1
$logStoreButton.Font = New-Object System.Drawing.Font("���C���I", 9)
$logStoreButton.Add_Click({
    # ���O�i�[�����������Ɏ���
    Write-Log "���O�i�[���������s���܂�" "INFO"
})
$form.Controls.Add($logStoreButton)
$script:logStoreButton = $logStoreButton

# ���O�\���G���A�i�����̃R���|�[�l���g���ړ��j
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 755)
$logTextBox.Size = New-Object System.Drawing.Size(780, 200)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.ReadOnly = $true
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logTextBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($logTextBox)
$script:logTextBox = $logTextBox

# �v���Z�X�R���g���[���̏�����
Update-ProcessControls

# �������b�Z�[�W
Write-Log "�A�v���P�[�V�������N�����܂���" "INFO"
Write-Log "�ݒ�t�@�C��: $configPath" "INFO"
Write-Log "�y�[�W��: $($script:pages.Count)" "INFO"

# �t�H�[����\��
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)
