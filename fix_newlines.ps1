# 改行を復元するスクリプト
$filePath = "Main.ps1"
$content = Get-Content -Path $filePath -Raw -Encoding UTF8

# コメントの後に変数や関数が続く場合に改行を追加
$content = $content -replace '(# [^\r\n]+)(\$[a-zA-Z])', '$1`r`n$2'
$content = $content -replace '(# [^\r\n]+)(function [a-zA-Z])', '$1`r`n$2'
$content = $content -replace '(# [^\r\n]+)(if \()', '$1`r`n$2'
$content = $content -replace '(# [^\r\n]+)(    if \()', '$1`r`n$2'
$content = $content -replace '(# [^\r\n]+)(Write-Host)', '$1`r`n$2'
$content = $content = $content -replace '(# [^\r\n]+)(        exit)', '$1`r`n$2'
$content = $content -replace '(# [^\r\n]+)(            return)', '$1`r`n$2'

# UTF-8 BOM付きで保存
$utf8WithBom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText((Resolve-Path $filePath), $content, $utf8WithBom)

Write-Host "改行を復元しました: $filePath"
