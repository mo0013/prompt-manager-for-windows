@echo off
setlocal enabledelayedexpansion

echo �X�^�[�g�A�b�v����V���[�g�J�b�g���폜���Ă��܂�...

set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=PromptManager.lnk
set POLICY_FILE=%~dp0powershell_policy_before_change.txt

if exist "%POLICY_FILE%" (
    for /f "tokens=*" %%i in (%POLICY_FILE%) do set previous_policy=%%i
    echo �ȑO�̎��s�|���V�[�𕜌����Ă��܂�: !previous_policy!
    powershell.exe -Command "Set-ExecutionPolicy '!previous_policy!' -Scope CurrentUser -Force"
) else (
    echo �|���V�[�t�@�C�������݂��Ȃ����߁A���s�|���V�[��ύX�ł��܂���ł����B
    for /f "tokens=*" %%i in ('powershell.exe -Command "Get-ExecutionPolicy -Scope CurrentUser"') do set current_policy=%%i
    echo ���݂̎��s�|���V�[: !current_policy!
)

powershell.exe -Command "Test-Path '%STARTUP_FOLDER%\%SHORTCUT_NAME%'"

if %ERRORLEVEL% == 0 (
    echo �V���[�g�J�b�g�����݂��܂��B�폜���܂����H (y/n)
    choice /c yn /m "�I��: "
    if errorlevel 2 goto :eof
    powershell.exe -Command "Remove-Item -Path '%STARTUP_FOLDER%\%SHORTCUT_NAME%' -Force"
    echo �X�^�[�g�A�b�v�V���[�g�J�b�g���폜����܂����B
) else (
    echo �X�^�[�g�A�b�v�V���[�g�J�b�g�͑��݂��Ȃ����߁A�I�����܂��B
)

pause