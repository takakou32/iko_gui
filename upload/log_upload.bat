@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo 使い方: %~nx0 [コピー元フォルダ] [コピー先パス]
    echo.
    echo 例: %~nx0 "C:\work\logs" "D:\backup"
    echo     → D:\backup\logs にコピーされます
    echo.
    echo 例: %~nx0 "C:\work\logs" "\\192.168.1.100\share\backup"
    echo     → \\192.168.1.100\share\backup\logs にコピーされます
    echo.
    pause
    exit /b 1
)

if "%~2"=="" (
    echo エラー: コピー先パスが指定されていません。
    pause
    exit /b 1
)

set SOURCE_DIR=%~1
set DEST_PATH=%~2

rem 末尾の\を削除
if "%SOURCE_DIR:~-1%"=="\" set SOURCE_DIR=%SOURCE_DIR:~0,-1%
if "%DEST_PATH:~-1%"=="\" set DEST_PATH=%DEST_PATH:~0,-1%

if not exist "%SOURCE_DIR%" (
    echo エラー: コピー元フォルダが存在しません: %SOURCE_DIR%
    pause
    exit /b 1
)

rem フォルダ名を取得
for %%F in ("%SOURCE_DIR%") do set FOLDER_NAME=%%~nxF

rem コピー先の完全パス
set DEST_FULL=%DEST_PATH%\%FOLDER_NAME%

echo ----------------------------------------
echo コピー元: %SOURCE_DIR%
echo コピー先: %DEST_FULL%
echo ----------------------------------------
echo.

robocopy "%SOURCE_DIR%" "%DEST_FULL%" /E /COPY:DAT /R:3 /W:5 /MT:8 /NP

set RESULT=!ERRORLEVEL!

echo.
if !RESULT! LSS 8 (
    echo ========================================
    echo コピーが完了しました！
    echo ========================================
) else (
    echo ========================================
    echo エラーが発生しました。
    echo エラーコード: !RESULT!
    echo ========================================
)

echo.
pause