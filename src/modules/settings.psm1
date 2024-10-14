<#
.SYNOPSIS
    設定管理モジュール
.DESCRIPTION
    このモジュールは、アプリケーションの設定を管理します。
    settings.xmlファイルから設定を読み込み、他のモジュールに提供します。
#>

$script:settingsPath = Join-Path $PSScriptRoot "..\..\config\settings.xml"
$script:settings = $null

function Initialize-Settings {
    if (Test-Path $script:settingsPath) {
        $script:settings = [xml](Get-Content $script:settingsPath)
    } else {
        Write-Error "設定ファイルが見つかりません: $script:settingsPath"
    }
}

function Get-Setting($xpath) {
    if ($null -eq $script:settings) {
        Initialize-Settings
    }
    return $script:settings.SelectSingleNode($xpath).InnerText
}

function Get-ApiKey($provider) {
    return Get-Setting "/Settings/ApiKeys/$provider"
}

function Set-ApiKey($provider, $key) {
    $node = $script:settings.SelectSingleNode("/Settings/ApiKeys/$provider")
    if ($node) {
        $node.InnerText = $key
        $script:settings.Save($script:settingsPath)
        return $true
    }
    return $false
}

Export-ModuleMember -Function Get-Setting, Initialize-Settings, Get-ApiKey, Set-ApiKey