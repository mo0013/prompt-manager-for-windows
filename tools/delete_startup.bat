@echo off
chcp 65001
setlocal enabledelayedexpansion

echo スタートアップからショートカットを削除しています...

set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=PromptManager.lnk
set POLICY_FILE=%~dp0powershell_policy_before_change.txt

if exist "%POLICY_FILE%" (
    for /f "tokens=*" %%i in (%POLICY_FILE%) do set previous_policy=%%i
    echo 以前の実行ポリシーを復元しています: !previous_policy!
    powershell.exe -Command "Set-ExecutionPolicy '!previous_policy!' -Scope CurrentUser -Force"
) else (
    echo ポリシーファイルが存在しないため、実行ポリシーを変更できませんでした。
    for /f "tokens=*" %%i in ('powershell.exe -Command "Get-ExecutionPolicy -Scope CurrentUser"') do set current_policy=%%i
    echo 現在の実行ポリシー: !current_policy!
)

powershell.exe -Command "Test-Path '%STARTUP_FOLDER%\%SHORTCUT_NAME%'"

if %ERRORLEVEL% == 0 (
    echo ショートカットが存在します。削除しますか？ (y/n)
    choice /c yn /m "選択: "
    if errorlevel 2 goto :eof
    powershell.exe -Command "Remove-Item -Path '%STARTUP_FOLDER%\%SHORTCUT_NAME%' -Force"
    echo スタートアップショートカットが削除されました。
) else (
    echo スタートアップショートカットは存在しないため、終了します。
)

pause