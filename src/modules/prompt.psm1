<#
.SYNOPSIS
    プロンプト管理アプリケーションのプロンプト処理モジュール
.DESCRIPTION
    このモジュールは、プロンプトファイルの読み込み、解析、保存を行います。
    Markdownファイルからプロンプトデータを抽出し、構造化されたデータとして提供します。
    また、新規プロンプトの作成と保存機能も提供します。
.NOTES
    このモジュールは、"data" フォルダ内のMarkdownファイルを対象としています。
    BOM無しUTF-8エンコーディングを使用してファイルの読み書きを行います。
#>

function Get-Prompts {
    param (
        [Parameter(Mandatory=$false)]
        [string]$Encoding = 'UTF8' 
    )
    <#
    .SYNOPSIS
        すべてのプロンプトを取得します。
    .DESCRIPTION
        "data" フォルダ内のすべてのMarkdownファイルを再帰的に読み込み、
        プロンプトデータを抽出して返します。
    .PARAMETER Encoding
        ファイルの読み込みに使用するエンコーディング。デフォルトは 'UTF8' です。
    .OUTPUTS
        System.Collections.ArrayList
    #>
    $prompts = New-Object System.Collections.ArrayList

    # "data" フォルダ内のすべての .md ファイルを取得
    $files = Get-ChildItem -Path "data" -Filter "*.md" -Recurse
    
    foreach ($file in $files) {
        $content = Get-Content -Path $file.FullName -Raw -Encoding $Encoding
        $relativePath = $file.FullName.Substring($file.FullName.IndexOf("data") + 5)
        $promptData = ConvertFrom-PromptContent -Content $content -FileName $relativePath
        $promptData | Add-Member -MemberType NoteProperty -Name "FilePath" -Value $file.FullName
        $promptData | Add-Member -MemberType NoteProperty -Name "Category" -Value $file.Directory.Name
        $prompts.Add($promptData) | Out-Null
    }

    return $prompts
}

function ConvertFrom-PromptContent {
    <#
    .SYNOPSIS
        Markdownファイルの内容をパースしてプロンプトデータを抽出します。
    .DESCRIPTION
        Markdownファイルの内容を解析し、タイトルと本文を抽出します。
        タイトルは最初の行（# で始まる）、本文は2行目以降とします。
    .PARAMETER Content
        Markdownファイルの内容
    .PARAMETER FileName
        Markdownファイルのファイル名（相対パス）
    .OUTPUTS
        System.Management.Automation.PSObject
    #>
    param (
        [string]$Content,
        [string]$FileName
    )

    $lines = $Content -split "`n"
    $title = $lines[0].Trim('#', ' ')
    $body = ($lines | Select-Object -Skip 1) -join "`n"

    return [PSCustomObject]@{
        Title = $title
        Content = $body.Trim()
        FileName = $FileName
    }
}

function Save-Prompt {
    <#
    .SYNOPSIS
        既存のプロンプトを保存します。
    .DESCRIPTION
        指定されたプロンプトオブジェクトを元のファイルパスに保存します。
        ディレクトリが存在しない場合は作成します。
    .PARAMETER Prompt
        保存するプロンプトオブジェクト
    #>
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Prompt
    )
    $dataFolder = "data"
    $filePath = Join-Path -Path $dataFolder -ChildPath $Prompt.FileName

    # ファイルのディレクトリが存在しない場合は作成
    $directory = Split-Path -Parent $filePath
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    # BOM無しUTF-8エンコーディングを指定
    $content = "# $($Prompt.Title)`n`n$($Prompt.Content)"
    [System.Text.Encoding]::UTF8.GetBytes($content) | Set-Content -Path $filePath -Encoding Byte
}

function Save-NewPrompt {
    <#
    .SYNOPSIS
        新規プロンプトを保存します。
    .DESCRIPTION
        指定されたプロンプトオブジェクトを新しいファイルとして保存します。
        カテゴリフォルダが存在しない場合は作成します。
    .PARAMETER Prompt
        保存する新規プロンプトオブジェクト
    .OUTPUTS
        成功時は保存されたプロンプトオブジェクト、失敗時は $null
    #>
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Prompt
    )

    $dataFolder = "data"
    $categoryFolder = Join-Path $dataFolder $Prompt.Category

    if (-not (Test-Path $categoryFolder)) {
        New-Item -ItemType Directory -Path $categoryFolder | Out-Null
    }

    $filePath = Join-Path $categoryFolder $Prompt.FileName
    $content = "# $($Prompt.Title)`n`n$($Prompt.Content)"
    
    # BOM無しUTF-8エンコーディングを使用
    [System.Text.Encoding]::UTF8.GetBytes($content) | Set-Content -Path $filePath -Encoding Byte
    # 保存が成功した場合、プロンプトオブジェクトを返す
    if (Test-Path $filePath) {
        return $Prompt
    } else {
        return $null
    }
}

function Remove-Prompt {
    <#
    .SYNOPSIS
        指定されたプロンプトを削除します。
    .DESCRIPTION
        指定されたプロンプトオブジェクトに対応するファイルを削除します。
    .PARAMETER Prompt
        削除するプロンプトオブジェクト
    .OUTPUTS
        [bool] 削除が成功したかどうか
    #>
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Prompt
    )

    try {
        Remove-Item -Path $Prompt.FilePath -Force
        return $true
    }
    catch {
        Write-Error "プロンプトの削除に失敗しました: $_"
        return $false
    }
}

# Export-ModuleMember に Save-NewPrompt を追加
Export-ModuleMember -Function Get-Prompts, Save-Prompt, Save-NewPrompt, Remove-Prompt