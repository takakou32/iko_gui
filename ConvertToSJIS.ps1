# ファイルをSJISエンコーディングに変換するスクリプト
# このスクリプトを実行して、.ps1と.batファイルをSJISに変換してください

$files = @("Main.ps1", "sample.bat", "sample2.bat")

foreach ($file in $files) {
    $filePath = Join-Path $PSScriptRoot $file
    
    if (Test-Path $filePath) {
        Write-Host "変換中: $file"
        
        # ファイルを読み込む（UTF-8として）
        $content = Get-Content $filePath -Raw -Encoding UTF8
        
        # SJISエンコーディングで保存
        $sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
        [System.IO.File]::WriteAllText($filePath, $content, $sjis)
        
        Write-Host "完了: $file" -ForegroundColor Green
    } else {
        Write-Host "ファイルが見つかりません: $file" -ForegroundColor Yellow
    }
}

Write-Host "`nすべてのファイルの変換が完了しました。" -ForegroundColor Green
