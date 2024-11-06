<#
.SYNOPSIS
    プロンプト管理アプリケーションのエントリポイント
.DESCRIPTION
    このスクリプトは、プロンプト管理アプリケーションのメインエントリポイントです。
    UIモジュールとプロンプトモジュールをロードし、システムトレイアイコンを初期化して
    アプリケーションを実行します。
.PARAMETER なし
.EXAMPLE
    .\main.ps1
.NOTES
    このスクリプトは、PowerShell 5.1以上で動作します。
    "data"フォルダ内のMarkdownファイルをプロンプトデータとして使用します。
#>

# スクリプトのディレクトリを取得
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# モジュール読み込みのエラー処理を追加
try {
    Import-Module -Name (Join-Path $scriptPath "modules\settings.psm1") -Force
    Import-Module -Name (Join-Path $scriptPath "modules\ui.psm1") -Force
    Import-Module -Name (Join-Path $scriptPath "modules\prompt.psm1") -Force
    Import-Module -Name (Join-Path $scriptPath "modules\llm.psm1") -Force

    # 設定を初期化
    Initialize-Settings
} catch {
    Write-Error "モジュールの読み込みまたは設定の初期化に失敗しました: $_"
    exit
}

# システムトレイアイコンを初期化
Initialize-TrayIcon

# アプリケーションを実行
[System.Windows.Forms.Application]::Run()