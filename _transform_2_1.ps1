$filePath = 'c:\work\iko_gui\UIPageRenderer.ps1'
$content = [System.IO.File]::ReadAllLines($filePath, [System.Text.Encoding]::UTF8)

# ブロック本体を抽出 (indices 337..677 = lines 338-678)
$blockBody = $content[337..677]

# 12スペース分デデント (16→4スペース)
$dedentedBody = $blockBody | ForEach-Object {
    if ($_.Length -ge 12 -and $_.Substring(0, 12) -eq '            ') {
        $_.Substring(12)
    } else {
        $_
    }
}

# 新関数ヘッダー
$funcHeader = @('', '# --- ページ1・2行レイアウト描画関数 ---', 'function Render-Page1And2Row {', '    param($i, $isPage1, $processConfig, $isEnabled)')
# 新関数フッター
$funcFooter = @('}')

# 新関数全体
$newFunction = $funcHeader + $dedentedBody + $funcFooter

# if ($useDrawioLayout) ブロックを関数呼び出しに置換
$replacement = @('            if ($useDrawioLayout) {', '                Render-Page1And2Row -i $i -isPage1 $isPage1 -processConfig $processConfig -isEnabled $isEnabled', '            }')

# 新コンテンツを組み立て
$part1 = $content[0..208]
$part2 = $content[209..335]
$part3 = $content[679..($content.Length - 1)]

$newContent = $part1 + $newFunction + $part2 + $replacement + $part3

Write-Host "元の行数: $($content.Length)"
Write-Host "新しい行数: $($newContent.Length)"

# UTF-8 BOM 付きで書き戻し
$encoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($filePath, $newContent, $encoding)

Write-Host "完了"
