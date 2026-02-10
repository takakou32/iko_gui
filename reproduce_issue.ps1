

try {
    Write-Host "Verifying Functions.ps1 syntax..."
    . c:\work\iko_gui\Functions.ps1
    Write-Host "Functions.ps1 loaded successfully."
    
    # Check if Update-PagePaths exists
    if (Get-Command Update-PagePaths -ErrorAction SilentlyContinue) {
        Write-Host "Update-PagePaths function exists."
    }
    else {
        Write-Error "Update-PagePaths function NOT found."
    }
    
    # Check if Get-CommonBasePath exists
    if (Get-Command Get-CommonBasePath -ErrorAction SilentlyContinue) {
        Write-Host "Get-CommonBasePath function exists."
    }
    else {
        Write-Error "Get-CommonBasePath function NOT found."
    }

    # Check if Save-ProcessDestinationPath exists
    if (Get-Command Save-ProcessDestinationPath -ErrorAction SilentlyContinue) {
        Write-Host "Save-ProcessDestinationPath function exists."
    }
    else {
        Write-Error "Save-ProcessDestinationPath function NOT found."
    }

}
catch {
    Write-Error "Syntax Check Failed: $($_.Exception.Message)"
}

