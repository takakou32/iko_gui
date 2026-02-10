
try {
    $script:configPath = "c:\work\iko_gui\config.json"
    $currentConfig = Get-Content $script:configPath -Encoding UTF8 | ConvertFrom-Json
    
    Write-Host "Config loaded."
    Write-Host "GlobalLogPath current value: '$($currentConfig.GlobalLogPath)'"
    
    # Simulate the logic used in UILayout.ps1
    $selectedPath = "C:\Test\LogPath"
    $PSScriptRoot = "c:\work\iko_gui"
    
    $savePath = $selectedPath
    if ($selectedPath.StartsWith($PSScriptRoot)) {
        $relativePath = $selectedPath.Substring($PSScriptRoot.Length).TrimStart("\")
        $savePath = $relativePath
    }
    
    Write-Host "Save Path: $savePath"

    # The problematic check
    $prop = $currentConfig.PSObject.Properties['GlobalLogPath']
    Write-Host "Property check result: '$prop'"
    
    if ($prop) {
        Write-Host "Property found, updating..."
        $currentConfig.GlobalLogPath = $savePath
    }
    else {
        Write-Host "Property NOT found, adding..."
        $currentConfig | Add-Member -MemberType NoteProperty -Name "GlobalLogPath" -Value $savePath
    }
    
    Write-Host "New GlobalLogPath value: '$($currentConfig.GlobalLogPath)'"
    
}
catch {
    Write-Error "Error occurred: $($_.Exception.Message)"
}
