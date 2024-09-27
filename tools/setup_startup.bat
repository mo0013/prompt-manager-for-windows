@echo off
setlocal enabledelayedexpansion

echo 現在の実行ポリシーを確認しています...
for /f "tokens=*" %%i in ('powershell.exe -Command "Get-ExecutionPolicy -Scope CurrentUser"') do set current_policy=%%i

if /i "%current_policy%" neq "RemoteSigned" (
    echo %current_policy% > "%~dp0powershell_policy_before_change.txt"
)

if /i "%current_policy%"=="Restricted" (
    echo 現在の実行ポリシーが Restricted です。RemoteSigned に変更します...
    powershell.exe -Command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
) else if /i "%current_policy%"=="AllSigned" (
    echo 現在の実行ポリシーが AllSigned です。RemoteSigned に変更します...
    powershell.exe -Command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
) else (
    echo 現在の実行ポリシー: %current_policy% - 変更は不要です。
)

echo スタートアップにショートカットを作成しています...
set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=PromptManager.lnk

REM 親ディレクトリのパスを取得
for %%I in ("%~dp0..") do set "PARENT_DIR=%%~fI"

powershell.exe -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%STARTUP_FOLDER%\%SHORTCUT_NAME%'); $s.TargetPath = 'powershell.exe'; $s.Arguments = '-WindowStyle Hidden -ExecutionPolicy Bypass -File ""!PARENT_DIR!\src\main.ps1""'; $s.WindowStyle = 7; $s.WorkingDirectory = '!PARENT_DIR!'; $s.Save()"

echo セットアップが完了しました。
echo スタートアップショートカットが以下の場所に作成されました:
echo %STARTUP_FOLDER%\%SHORTCUT_NAME%
pause