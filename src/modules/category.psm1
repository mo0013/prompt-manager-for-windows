<#
.SYNOPSIS
    カテゴリ管理モジュール
.DESCRIPTION
    このモジュールは、カテゴリの管理に関連する機能を提供します。
    カテゴリの一覧取得、名前変更、削除などの操作を行います。
#>

function Get-CategoryList {
    <#
    .SYNOPSIS
        カテゴリの一覧を取得します。
    .DESCRIPTION
        dataフォルダ内のサブフォルダ名をカテゴリとして取得します。
    #>
    $dataFolder = "data"
    $dataPath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\$dataFolder"
    Write-Host "データパス: $dataPath"
    $categories = Get-ChildItem -Path $dataPath -Directory | Select-Object -ExpandProperty Name
    Write-Host "取得されたカテゴリ: $($categories -join ', ')"
    return $categories
}

function Rename-Category {
    param (
        [string]$OldName,
        [string]$NewName
    )

    $dataFolder = "data"
    $dataPath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\$dataFolder"
    $oldPath = Join-Path -Path $dataPath -ChildPath $OldName
    $newPath = Join-Path -Path $dataPath -ChildPath $NewName

    if (Test-Path $newPath) {
        throw "指定された新しい名前のカテゴリは既に存在します。"
    }

    try {
        Rename-Item -Path $oldPath -NewName $NewName -ErrorAction Stop
        return $true
    } catch {
        throw "カテゴリ名の変更中にエラーが発生しました: $_"
    }
}

function Update-CategoryList {
    <#
    .SYNOPSIS
        カテゴリリストを更新します。
    .DESCRIPTION
        指定されたListBoxコントロールにカテゴリ一覧を表示します。
    .PARAMETER ListBox
        更新するListBoxコントロール
    #>
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ListBox]$ListBox
    )

    $ListBox.Items.Clear()
    $categories = Get-CategoryList
    Write-Host "更新されるカテゴリ: $($categories -join ', ')"
    foreach ($category in $categories) {
        $ListBox.Items.Add($category)
    }
}

# 他のカテゴリ関連の関数をここに追加

Export-ModuleMember -Function Get-CategoryList, Rename-Category, Update-CategoryList