@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo ログフォルダをアップロード
echo ========================================
echo.

set SOURCE_DIR=C:\work\iko_gui\logs
set DEST_PATH=C:\work\iko_gui\summary

rem 末尾の\を削除
if "%SOURCE_DIR:~-1%"=="\" set SOURCE_DIR=%SOURCE_DIR:~0,-1%
if "%DEST_PATH:~-1%"=="\" set DEST_PATH=%DEST_PATH:~0,-1%

if not exist "%SOURCE_DIR%" (
    echo エラー: コピー元フォルダが存在しません: %SOURCE_DIR%
    pause
    exit /b 1
)

rem コピー先パスが存在しない場合は作成
if not exist "%DEST_PATH%" (
    echo コピー先パスを作成します: %DEST_PATH%
    mkdir "%DEST_PATH%"
    if !ERRORLEVEL! NEQ 0 (
        echo エラー: コピー先パスの作成に失敗しました。
        pause
        exit /b 1
    )
)

rem フォルダ名を取得
for %%F in ("%SOURCE_DIR%") do set FOLDER_NAME=%%~nxF

rem コピー先の完全パス
set DEST_FULL=%DEST_PATH%\%FOLDER_NAME%

echo コピー元: %SOURCE_DIR%
echo コピー先: %DEST_FULL%
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