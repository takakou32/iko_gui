@echo off
chcp 932 > nul
echo サンプルバッチファイル2を実行中...
echo 現在の日時: %date% %time%
timeout /t 2 /nobreak > nul
echo サンプルバッチファイル2の実行が完了しました。
exit /b 0
