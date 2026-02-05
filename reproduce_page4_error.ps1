

# Reproduce Page 4 Errors

# Mock WinForms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Mock Global Variables
$script:pages = @(
    @{}, @{}, @{}, # 0, 1, 2
    @{ # Page 4 (Index 3)
        JsonPath  = "c:\work\iko_gui\page4.json"
        Processes = @() # Will be loaded by Get-CurrentPageProcesses
    }
)
$script:currentPage = 3
$script:processesPerPage = 4
$script:editMode = $false
$script:logDir = "c:\work\iko_gui\logs"
$PSScriptRoot = "c:\work\iko_gui"

# Mock UI Elements
$script:processPanel = New-Object System.Windows.Forms.Panel
$script:processControls = @()
$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Controls.Add($script:processPanel)
$script:pageLabel = New-Object System.Windows.Forms.Label
$script:titleLabel = New-Object System.Windows.Forms.Label
$script:addRowButton = New-Object System.Windows.Forms.Button
$script:deleteRowButton = New-Object System.Windows.Forms.Button
$script:logStoragePathTextBox = New-Object System.Windows.Forms.TextBox
$script:logStoragePathTextBox2 = New-Object System.Windows.Forms.TextBox

# Source Functions (ignoring unrelated errors)
. c:\work\iko_gui\Functions.ps1

# Mock Logger that prints to console (override Functions.ps1)
function Write-Log {
    param($msg, $level)
    Write-Host "[$level] $msg"
}

# Load Page 4 JSON logic (Mocking Get-CurrentPageProcesses to use actual file or raw json)
# We rely on Functions.ps1's Get-CurrentPageProcesses. 
# Ensure page4.json exists or we mock the return
if (-not (Test-Path "c:\work\iko_gui\page4.json")) {
    Write-Host "page4.json not found, using mock data"
    $script:pages[3].Processes = @(
        @{ Name = "P1"; KdlSourcePath = "C:\"; KdlDestPath = "C:\" },
        @{ Name = "P2"; KdlSourcePath = "C:\"; KdlDestPath = "C:\" },
        @{ Name = "P3"; V1CsvDestPath = "C:\" },
        @{ Name = "New"; BatchFiles = @(@{Path = "" }) }
    )
}

Write-Host "Starting Update-ProcessControls..."
try {
    Update-ProcessControls
    Write-Host "Update-ProcessControls execution finished."
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Host "$($_.ScriptStackTrace)"
}
