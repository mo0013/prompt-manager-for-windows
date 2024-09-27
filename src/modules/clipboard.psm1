<#
.SYNOPSIS
    プロンプト管理アプリケーションのクリップボード操作モジュール
.DESCRIPTION
    このモジュールは、クリップボードへのテキストのコピーと取得を管理します。
    UTF-8エンコーディングを使用してクリップボード操作を行います。
.NOTES
    このモジュールは、System.Windows.Formsを使用してクリップボード操作を行います。
    クリップボードの操作に失敗した場合、エラーメッセージを返します。
#>

Add-Type -AssemblyName System.Windows.Forms

function Copy-TextToClipboard {
    <#
    .SYNOPSIS
        指定されたテキストをクリップボードにコピーします。
    .DESCRIPTION
        指定されたテキストをUnicodeテキストとしてクリップボードにコピーし、操作の結果を返します。
        コピーに成功した場合は成功メッセージを、失敗した場合はエラーメッセージを含むオブジェクトを返します。
    .PARAMETER Text
        クリップボードにコピーするテキスト
    .OUTPUTS
        System.Management.Automation.PSObject
        Success (bool): 操作が成功したかどうか
        Message (string): 操作の結果メッセージ
    .EXAMPLE
        Copy-TextToClipboard -Text "これはテストです。"
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    try {
        [System.Windows.Forms.Clipboard]::SetText($Text, [System.Windows.Forms.TextDataFormat]::UnicodeText)
        return [PSCustomObject]@{
            Success = $true
            Message = "テキストがクリップボードにコピーされました。"
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Message = "クリップボードへのコピーに失敗しました: $_"
        }
    }
}

function Get-ClipboardText {
    <#
    .SYNOPSIS
        クリップボードからテキストを取得します。
    .DESCRIPTION
        クリップボードの現在の内容をUnicodeテキストとして取得し、返します。
        クリップボードからの取得に失敗した場合、エラーメッセージを出力し、$nullを返します。
    .OUTPUTS
        [string] クリップボードの内容、または失敗時は $null
    .EXAMPLE
        $clipboardContent = Get-ClipboardText
    #>
    try {
        return [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::UnicodeText)
    }
    catch {
        Write-Error "クリップボードからのテキスト取得に失敗しました: $_"
        return $null
    }
}

Export-ModuleMember -Function Copy-TextToClipboard, Get-ClipboardText