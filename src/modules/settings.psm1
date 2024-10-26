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
        try {
            $script:settings = [xml](Get-Content $script:settingsPath)
        } catch {
            Write-Error "設定ファイルの読み込み中にエラーが発生しました: $_"
            $script:settings = $null
        }
    } else {
        Write-Error "設定ファイルが見つかりません: $script:settingsPath"
    }
}

function Get-Setting($xpath) {
    <#
    .SYNOPSIS
        指定されたXPathに基づいて設定値を取得します。
    .DESCRIPTION
        設定ファイルから指定されたXPathに対応する設定値を取得します。
        設定が見つからない場合はエラーを出力し、nullを返します。
    .PARAMETER xpath
        取得したい設定のXPath
    .OUTPUTS
        設定値（文字列）またはnull
    #>
    if ($null -eq $script:settings) {
        Initialize-Settings
    }
    $node = $script:settings.SelectSingleNode($xpath)
    if ($node) {
        return $node.InnerText
    } else {
        Write-Error "設定が見つかりません: $xpath"
        return $null
    }
}

function Get-ApiKey($provider) {
    return Get-Setting "/Settings/ApiKeys/$provider"
}

function Set-ApiKey($provider, $key) {
    <#
    .SYNOPSIS
        指定されたプロバイダーのAPIキーを設定します。
    .DESCRIPTION
        設定ファイル内の指定されたプロバイダーのAPIキーを更新します。
        設定が初期化されていない場合は、初期化を行います。
    .PARAMETER provider
        APIキーを設定するプロバイダー名
    .PARAMETER key
        設定するAPIキー
    .OUTPUTS
        [bool] 設定の更新が成功したかどうか
    #>
    if ($null -eq $script:settings) {
        Initialize-Settings
    }
    $node = $script:settings.SelectSingleNode("/Settings/ApiKeys/$provider")
    if ($node) {
        $node.InnerText = $key
        $script:settings.Save($script:settingsPath)
        return $true
    }
    return $false
}

Export-ModuleMember -Function Get-Setting, Initialize-Settings, Get-ApiKey, Set-ApiKey