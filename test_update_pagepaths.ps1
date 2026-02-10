
# Test Update-PagePaths
$PSScriptRoot = "c:\work\iko_gui"
$script:configPath = Join-Path $PSScriptRoot "config\json\config.json"

# Mock GUI objects
$script:processPanel = New-Object System.Windows.Forms.Panel
$script:logStoragePathTextBox = New-Object System.Windows.Forms.TextBox

# Load Functions
. "$PSScriptRoot\Functions.ps1"

# Load Config
if (Test-Path $script:configPath) {
    Write-Host "Loading config from $script:configPath"
    $script:config = Get-Content $script:configPath -Encoding UTF8 | ConvertFrom-Json
    $script:pages = $script:config.Pages
}
else {
    Write-Host "Config not found!"
    exit
}

# Test Page 1 Loading
$script:currentPage = 0
Write-Host "Testing Page 0 (Page 1)..."
$pageConfig = $script:pages[0]
Write-Host "JsonPath in memory: $($pageConfig.JsonPath)"

# Call Update-PagePaths (which logs errors)
# We need to mock Write-Log to see output
function Write-Log {
    param($msg, $type, $idx)
    Write-Host "[$type] $msg"
}

Update-PagePaths

Write-Host "Done."
