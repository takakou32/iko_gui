@echo off
chcp 65001 >nul

echo ========================================
echo ログフォルダをアップロード
echo ========================================
echo.

rem このバッチファイルと同じフォルダにあるcopy_folder.batを呼び出す
call log_upload.bat C:\work\iko_gui\logs C:\work\iko_gui\summary
pause
exit /b %ERRORLEVEL%