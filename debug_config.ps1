
# Debug Configuration Loading
$PSScriptRoot = "c:\work\iko_gui"
$configPath = Join-Path $PSScriptRoot "config\json\config.json"

Write-Host "Config Path: $configPath"
if (Test-Path $configPath) {
    Write-Host "Config file exists."
    try {
        $config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json
        Write-Host "Config loaded successfully."
        
        if ($config.Pages) {
            foreach ($page in $config.Pages) {
                Write-Host "Page: $($page.Title)"
                Write-Host "  Raw JsonPath: $($page.JsonPath)"
                $resolvedPath = Join-Path $PSScriptRoot $page.JsonPath
                Write-Host "  Resolved Path: $resolvedPath"
                if (Test-Path $resolvedPath) {
                    Write-Host "  [OK] File exists."
                }
                else {
                    Write-Host "  [ERROR] File NOT found."
                }
            }
        }
        else {
            Write-Host "No Pages defined in config."
        }
    }
    catch {
        Write-Host "Error loading json: $($_.Exception.Message)"
    }
}
else {
    Write-Host "Config file NOT found at $configPath"
    # Check root config
    $rootConfig = Join-Path $PSScriptRoot "config.json"
    if (Test-Path $rootConfig) {
        Write-Host "Found config.json at ROOT: $rootConfig"
    }
    else {
        Write-Host "No config.json at ROOT either."
    }
}
