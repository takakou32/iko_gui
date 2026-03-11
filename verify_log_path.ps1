# 検証スクリプト: Save-PagePaths のバリデーションと相対パス変換のテスト
. "$PSScriptRoot\Functions.ps1"

# テスト用のグローバル変数設定
$script:globalLogPath = "C:\work\fizz-buzz"
$script:currentPage = 2 # 3ページ目（インデックス2）
$script:pages = @(
    @{ Title = "Page 1"; JsonPath = "config\json\page1.json" },
    @{ Title = "Page 2"; JsonPath = "config\json\page2.json" },
    @{ Title = "Page 3"; JsonPath = "config\json\page3.json" }
)

Write-Host "--- Test 1: 共通パス内のフォルダを選択 ---"
$testPath1 = "C:\work\fizz-buzz\logs\test"
$result1 = Save-PagePaths -LogStoragePath $testPath1
Write-Host "Result: $result1"

Write-Host "`n--- Test 2: 共通パス外のフォルダを選択 ---"
$testPath2 = "C:\PerfLogs"
$result2 = Save-PagePaths -LogStoragePath $testPath2
Write-Host "Result: $result2 (Expected: False)"

Write-Host "`n--- Test 3: LogStoragePath2 (フルパス維持) の確認 ---"
$testPath3 = "\\network\share\logs"
$result3 = Save-PagePaths -LogStoragePath2 $testPath3
Write-Host "Result: $result3"

Write-Host "`n--- JSON内容の確認 ---"
Get-Content (Join-Path $PSScriptRoot "config\json\page3.json") -Raw | ConvertFrom-Json | Select-Object LogStoragePath, LogStoragePath2 | Format-List
