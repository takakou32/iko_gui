
# Load Functions
. "$PSScriptRoot\Functions.ps1"

# Mock variables
$script:currentPage = 0
$script:pages = @(
    @{
        LogStoragePath = "C:\work\test\fizz-buzz"
        JsonPath       = "page1.json"
        Processes      = @(
            @{ BatchFiles = @(@{Name = "Test"; Path = "" }) }
        )
    }
)
$script:config = @{ Title = "Test" }

# Test Case 1: Valid Path (Inside LogStoragePath)
$validPath = "C:\work\test\fizz-buzz\test.bat"
Write-Host "Testing Valid Path: $validPath"
$result = Save-BatchFilePath -ProcessIndex 0 -BatchFilePath $validPath -BatchIndex 0
Write-Host "Result (Should be True): $result"

# Test Case 2: Invalid Path (Outside LogStoragePath)
$invalidPath = "C:\work\SpringSample\gradlew.bat"
Write-Host "Testing Invalid Path: $invalidPath"
$result = Save-BatchFilePath -ProcessIndex 0 -BatchFilePath $invalidPath -BatchIndex 0
Write-Host "Result (Should be False): $result"
