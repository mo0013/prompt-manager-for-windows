@echo off
setlocal enabledelayedexpansion

echo ���݂̎��s�|���V�[���m�F���Ă��܂�...
for /f "tokens=*" %%i in ('powershell.exe -Command "Get-ExecutionPolicy -Scope CurrentUser"') do set current_policy=%%i

if /i "%current_policy%" neq "RemoteSigned" (
    echo %current_policy% > "%~dp0powershell_policy_before_change.txt"
)

if /i "%current_policy%"=="Restricted" (
    echo ���݂̎��s�|���V�[�� Restricted �ł��BRemoteSigned �ɕύX���܂�...
    powershell.exe -Command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
) else if /i "%current_policy%"=="AllSigned" (
    echo ���݂̎��s�|���V�[�� AllSigned �ł��BRemoteSigned �ɕύX���܂�...
    powershell.exe -Command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
) else (
    echo ���݂̎��s�|���V�[: %current_policy% - �ύX�͕s�v�ł��B
)

echo �X�^�[�g�A�b�v�ɃV���[�g�J�b�g���쐬���Ă��܂�...
set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=PromptManager.lnk

REM �e�f�B���N�g���̃p�X���擾
for %%I in ("%~dp0..") do set "PARENT_DIR=%%~fI"

powershell.exe -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%STARTUP_FOLDER%\%SHORTCUT_NAME%'); $s.TargetPath = 'powershell.exe'; $s.Arguments = '-WindowStyle Hidden -ExecutionPolicy Bypass -File ""!PARENT_DIR!\src\main.ps1""'; $s.WindowStyle = 7; $s.WorkingDirectory = '!PARENT_DIR!'; $s.Save()"

echo �Z�b�g�A�b�v���������܂����B
echo �X�^�[�g�A�b�v�V���[�g�J�b�g���ȉ��̏ꏊ�ɍ쐬����܂���:
echo %STARTUP_FOLDER%\%SHORTCUT_NAME%
pause