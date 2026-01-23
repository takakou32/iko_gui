@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo 使い方: %~nx0 [コピー元フォルダ] [コピー先の親フォルダ]
    exit /b 1
)

if "%~2"=="" (
    echo エラー: コピー先の親フォルダが指定されていません。
    exit /b 1
)

set SOURCE_DIR=%~1
if "%SOURCE_DIR:~-1%"=="\" set SOURCE_DIR=%SOURCE_DIR:~0,-1%

if not exist "%SOURCE_DIR%" (
    echo エラー: コピー元フォルダが存在しません: %SOURCE_DIR%
    exit /b 1
)

for %%F in ("%SOURCE_DIR%") do set FOLDER_NAME=%%~nxF

set DEST_DIR=%~2
if "%DEST_DIR:~-1%"=="\" set DEST_DIR=%DEST_DIR:~0,-1%

set DEST_FULL=%DEST_DIR%\%FOLDER_NAME%

echo コピー元: %SOURCE_DIR%
echo コピー先: %DEST_FULL%
echo.

robocopy "%SOURCE_DIR%" "%DEST_FULL%" /E /COPY:DAT /R:3 /W:5 /MT:8 /NP /NFL /NDL

if !ERRORLEVEL! LSS 8 (
    echo コピー完了
    exit /b 0
) else (
    echo エラー発生
    exit /b 1
)