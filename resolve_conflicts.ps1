# コンフリクト解決スクリプト
$filePath = "Main.ps1"
$content = Get-Content -Path $filePath -Raw -Encoding UTF8

# コンフリクトマーカーを削除してHEADバージョンを採用
$content = $content -replace '(?s)<<<<<<< HEAD\r?\n(.*?)\r?\n=======\r?\n.*?\r?\n>>>>>>> [^\r\n]+\r?\n', '$1'

# UTF-8 BOM付きで保存
$utf8WithBom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText((Resolve-Path $filePath), $content, $utf8WithBom)

Write-Host "コンフリクトを解決しました: $filePath"
