﻿<#
.SYNOPSIS
    プロンプト管理アプリケーションのユーザーインターフェースモジュール
.DESCRIPTION
    このモジュールは、プロンプト管理アプリケーションのユーザーインターフェースを提供します。
    メインウィンドウ、トレイアイコン、各種フォーム（新規作成、編集、プレビュー）の作成と
    管理を行います。また、プロンプトの表示、コピー、編集などの機能も提供します。
.NOTES
    このモジュールは、Windows Forms を使用してGUIを構築しています。
    System.Windows.Forms と System.Drawing アセンブリを使用しています。
#>

# スクリプトのパスを取得
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# クリップボードモジュールをインポート
Import-Module "$scriptPath\clipboard.psm1"
Import-Module "$scriptPath\llm.psm1"

# Windows Formsの名前空間をインポート
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# グローバル変数
$script:MainForm = $null
$script:PromptListBox = $null
$script:CopyButton = $null
$script:TrayIcon = $null
$script:IsExiting = $false

# グローバル変数にフィルタリングされたプロンプトリストを追加
$script:FilteredPrompts = @()

function Initialize-TrayIcon {
    <#
    .SYNOPSIS
        トレイアイコンを初期化します。
    .DESCRIPTION
        システムトレイにアイコンを作成し、コンテキストメニューを設定します。
        アイコンのクリックイベントも設定します。
    #>
    $script:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:TrayIcon.Text = "プロンプト管理アプリ"
    
    # アイコンファイルのパスを設定
    $iconPath = Join-Path $PSScriptRoot "..\..\assets\icon.ico"
    if (Test-Path $iconPath) {
        $script:TrayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    } else {
        Write-Warning "アイコンファイルが見つかりません: $iconPath"
    }

    # コンテキストメニューの作成
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $openItem = $contextMenu.Items.Add("開く")
    $openItem.Add_Click({ Show-PromptManagerMainWindow })
    $exitItem = $contextMenu.Items.Add("終了")
    $exitItem.Add_Click({ Exit-Application })

    $script:TrayIcon.ContextMenuStrip = $contextMenu

    # アイコンのクリックイベントを設定
    $script:TrayIcon.Add_Click({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Show-PromptManagerMainWindow
        }
    })

    $script:TrayIcon.Visible = $true
}

function Show-PromptManagerMainWindow {
    <#
    .SYNOPSIS
        プロンプト管理アプリのメインウィンドウを表示します。
    .DESCRIPTION
        メインフォームを作成し、各種コントロール（リストボックス、ボタンなど）を
        配置します。また、プロンプトの一覧を表示し、各種操作を行うためのUIを提供します。
    #>
    if ($null -eq $script:MainForm -or $script:MainForm.IsDisposed) {    
        # メインフォームの作成
        $script:MainForm = New-Object System.Windows.Forms.Form
        $script:MainForm.Text = "プロンプト管理アプリ"
        $script:MainForm.Size = New-Object System.Drawing.Size(500,370)
        $script:MainForm.StartPosition = "CenterScreen"

        # リストボックスの作成
        $script:PromptListBox = New-Object System.Windows.Forms.ListBox
        $script:PromptListBox.Location = New-Object System.Drawing.Point(10,10)
        $script:PromptListBox.Size = New-Object System.Drawing.Size(360,300) 
        $script:PromptListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                      [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                                      [System.Windows.Forms.AnchorStyles]::Left -bor 
                                      [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($script:PromptListBox)

        # プレビューボタンの作成
        $script:PreviewButton = New-Object System.Windows.Forms.Button
        $script:PreviewButton.Location = New-Object System.Drawing.Point(380,10)
        $script:PreviewButton.Size = New-Object System.Drawing.Size(100,30)
        $script:PreviewButton.Text = "プレビュー"
        $script:PreviewButton.Add_Click({ Show-PromptPreview })
        $script:PreviewButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                       [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($script:PreviewButton)

        # コピーボタンの作成
        $script:CopyButton = New-Object System.Windows.Forms.Button
        $script:CopyButton.Location = New-Object System.Drawing.Point(380,50)
        $script:CopyButton.Size = New-Object System.Drawing.Size(100,30)
        $script:CopyButton.Text = "コピー"
        $script:CopyButton.Add_Click({ Copy-SelectedPrompt })
        $script:CopyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                     [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($script:CopyButton)

        # 新規作成ボタンの作成
        $newButton = New-Object System.Windows.Forms.Button
        $newButton.Location = New-Object System.Drawing.Point(380,90)
        $newButton.Size = New-Object System.Drawing.Size(100,30)
        $newButton.Text = "新規作成"
        $newButton.Add_Click({ Show-NewPromptForm })
        $newButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                             [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($newButton)

        # 編集ボタンの作成
        $script:EditButton = New-Object System.Windows.Forms.Button
        $script:EditButton.Location = New-Object System.Drawing.Point(380,130)
        $script:EditButton.Size = New-Object System.Drawing.Size(100,30)
        $script:EditButton.Text = "編集"
        $script:EditButton.Add_Click({ Edit-SelectedPrompt })
        $script:EditButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                    [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($script:EditButton)

        # 削除ボタンの作成
        $script:DeleteButton = New-Object System.Windows.Forms.Button
        $script:DeleteButton.Location = New-Object System.Drawing.Point(380,170)
        $script:DeleteButton.Size = New-Object System.Drawing.Size(100,30)
        $script:DeleteButton.Text = "削除"
        $script:DeleteButton.Add_Click({ Remove-SelectedPrompt })
        $script:DeleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                      [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($script:DeleteButton)

        # カテゴリフィルタ用コンボボックスの作成
        $script:CategoryFilter = New-Object System.Windows.Forms.ComboBox
        $script:CategoryFilter.Location = New-Object System.Drawing.Point(10, 10)
        $script:CategoryFilter.Size = New-Object System.Drawing.Size(150, 20)
        $script:CategoryFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $script:CategoryFilter.Add_SelectedIndexChanged({ Update-PromptList })
        $script:MainForm.Controls.Add($script:CategoryFilter)
        $script:PromptListBox.Location = New-Object System.Drawing.Point(10, 40)
        $script:PromptListBox.Size = New-Object System.Drawing.Size(360, 280) 

        # プロンプトリストの更新
        Update-PromptList

        # カテゴリ管理ボタンの作成
        $categoryManageButton = New-Object System.Windows.Forms.Button
        $categoryManageButton.Location = New-Object System.Drawing.Point(380, 210)
        $categoryManageButton.Size = New-Object System.Drawing.Size(100, 30)
        $categoryManageButton.Text = "カテゴリ管理"
        $categoryManageButton.Add_Click({ Show-CategoryManagementForm })
        $categoryManageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                       [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($categoryManageButton)

        # API Key設定ボタンの作成
        $apiKeySettingsButton = New-Object System.Windows.Forms.Button
        $apiKeySettingsButton.Location = New-Object System.Drawing.Point(380, 250)
        $apiKeySettingsButton.Size = New-Object System.Drawing.Size(100, 30)
        $apiKeySettingsButton.Text = "API Key設定"
        $apiKeySettingsButton.Add_Click({ Show-ApiKeySettingsForm })
        $apiKeySettingsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                       [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($apiKeySettingsButton)

        # 終了ボタンの追加（1つ分の空間を空けて配置）
        $exitButton = New-Object System.Windows.Forms.Button
        $exitButton.Location = New-Object System.Drawing.Point(380, 290)
        $exitButton.Size = New-Object System.Drawing.Size(100, 30)
        $exitButton.Text = "終了"
        $exitButton.Add_Click({ 
            if ($null -ne $script:MainForm) {
                $script:MainForm.Hide()
            }
        })
        $exitButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                             [System.Windows.Forms.AnchorStyles]::Right
        $script:MainForm.Controls.Add($exitButton)

        # フォームが閉じられたときの処理
        $script:MainForm.Add_FormClosing({
            if ($_.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
                $_.Cancel = $true
                $script:MainForm.Hide()
            }
        })
        # メインフォームの Loaded イベントでカテゴリリストとプロンプトリストを更新
        $script:MainForm.Add_Load({
            Update-CategoryList
            Update-PromptList
        })
    }
    if ($null -ne $script:MainForm -and -not $script:MainForm.IsDisposed) {
        $script:MainForm.Show()
        $script:MainForm.Activate()
    } else {
        Write-Error "メインフォームの作成に失敗しました。"
    }
}

function Show-NewPromptForm {
    <#
    .SYNOPSIS
        新規プロンプト作成フォームを表示します。
    .DESCRIPTION
        新しいプロンプトを作成するためのフォームを表示します。
        タイトル、カテゴリ、内容を入力し、保存することができます。
    #>
    $newPromptForm = New-Object System.Windows.Forms.Form
    $newPromptForm.Text = "新規プロンプト作成"
    $newPromptForm.Size = New-Object System.Drawing.Size(500,400)
    $newPromptForm.StartPosition = "CenterScreen"

    # タイトル入力フィールド
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(10,20)
    $titleLabel.Size = New-Object System.Drawing.Size(100,20)
    $titleLabel.Text = "タイトル:"
    $newPromptForm.Controls.Add($titleLabel)

    $titleTextBox = New-Object System.Windows.Forms.TextBox
    $titleTextBox.Location = New-Object System.Drawing.Point(120,20)
    $titleTextBox.Size = New-Object System.Drawing.Size(350,20)
    $newPromptForm.Controls.Add($titleTextBox)

    # カテゴリ選択フィールド
    $categoryLabel = New-Object System.Windows.Forms.Label
    $categoryLabel.Location = New-Object System.Drawing.Point(10,50)
    $categoryLabel.Size = New-Object System.Drawing.Size(100,20)
    $categoryLabel.Text = "カテゴリ:"
    $newPromptForm.Controls.Add($categoryLabel)

    $categoryComboBox = New-Object System.Windows.Forms.ComboBox
    $categoryComboBox.Location = New-Object System.Drawing.Point(120,50)
    $categoryComboBox.Size = New-Object System.Drawing.Size(250,20)
    $categoryComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $categories = Get-ChildItem -Path "data" -Directory | Select-Object -ExpandProperty Name
    $categoryComboBox.Items.AddRange($categories)
    $newPromptForm.Controls.Add($categoryComboBox)

    # 新規カテゴリ追加ボタン
    $newCategoryButton = New-Object System.Windows.Forms.Button
    $newCategoryButton.Location = New-Object System.Drawing.Point(380,50)
    $newCategoryButton.Size = New-Object System.Drawing.Size(90,20)
    $newCategoryButton.Text = "新規カテゴリ"
    $newCategoryButton.Add_Click({
        $newCategory = [Microsoft.VisualBasic.Interaction]::InputBox("新しいカテゴリ名を入力してください", "新規カテゴリ")
        if ($newCategory) {
            $categoryComboBox.Items.Add($newCategory)
            $categoryComboBox.SelectedItem = $newCategory
        }
    })
    $newPromptForm.Controls.Add($newCategoryButton)

    # プロンプト内容入力フィールド
    $contentLabel = New-Object System.Windows.Forms.Label
    $contentLabel.Location = New-Object System.Drawing.Point(10,80)
    $contentLabel.Size = New-Object System.Drawing.Size(100,20)
    $contentLabel.Text = "内容:"
    $newPromptForm.Controls.Add($contentLabel)

    $contentTextBox = New-Object System.Windows.Forms.TextBox
    $contentTextBox.Location = New-Object System.Drawing.Point(10,100)
    $contentTextBox.Size = New-Object System.Drawing.Size(460,200)
    $contentTextBox.Multiline = $true
    $contentTextBox.ScrollBars = "Vertical"
    $contentTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                              [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                              [System.Windows.Forms.AnchorStyles]::Left -bor 
                              [System.Windows.Forms.AnchorStyles]::Right
    $newPromptForm.Controls.Add($contentTextBox)

    # 保存ボタン
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(300,320)
    $saveButton.Size = New-Object System.Drawing.Size(75,23)
    $saveButton.Text = "保存"
    $saveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                         [System.Windows.Forms.AnchorStyles]::Right
    $saveButton.Add_Click({
        if (-not $titleTextBox.Text -or -not $categoryComboBox.SelectedItem -or -not $contentTextBox.Text) {
            [System.Windows.Forms.MessageBox]::Show(
                "すべての項目を入力してください。", 
                "エラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
    
        $newPrompt = [PSCustomObject]@{
            Title = $titleTextBox.Text
            Category = $categoryComboBox.SelectedItem
            Content = $contentTextBox.Text
            FileName = "$($titleTextBox.Text -replace '[^\w\-\.]', '_').md"
        }
    
        $savedPrompt = Save-NewPrompt $newPrompt
        if ($null -ne $savedPrompt) {
            [System.Windows.Forms.MessageBox]::Show(
                "新しいプロンプトを保存しました。: $($savedPrompt)", 
                "保存完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            # カテゴリリストを更新
            Update-CategoryList
            # メインウィンドウのリストボックスを更新
            Update-PromptList
            
            $newPromptForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "プロンプトの保存に失敗しました。ファイル名を確認してください。", 
                "エラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    $newPromptForm.Controls.Add($saveButton)

    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(390,320)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "キャンセル"
    $cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                           [System.Windows.Forms.AnchorStyles]::Right
    $cancelButton.Add_Click({ $newPromptForm.Close() })
    $newPromptForm.Controls.Add($cancelButton)

    $newPromptForm.ShowDialog()
}

function Edit-SelectedPrompt {
    <#
    .SYNOPSIS
        選択されたプロンプトを編集します。
    .DESCRIPTION
        リストボックスで選択されたプロンプトの編集フォームを表示し、
        内容とカテゴリを編集して保存することができます。
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = $script:FilteredPrompts[$selectedIndex]
        $editForm = New-Object System.Windows.Forms.Form
        $editForm.Text = "プロンプト編集"
        $editForm.Size = New-Object System.Drawing.Size(600, 500)  # フォームのサイズを大きくしました
        $editForm.StartPosition = "CenterScreen"

        # 現在のカテゴリを表示するラベル
        $currentCategoryLabel = New-Object System.Windows.Forms.Label
        $currentCategoryLabel.Location = New-Object System.Drawing.Point(10, 10)
        $currentCategoryLabel.Size = New-Object System.Drawing.Size(200, 20)
        $currentCategoryLabel.Text = "現在のカテゴリ: $($selectedPrompt.Category)"
        $editForm.Controls.Add($currentCategoryLabel)

        # カテゴリ選択用ラジオボタン
        $existingCategoryRadio = New-Object System.Windows.Forms.RadioButton
        $existingCategoryRadio.Location = New-Object System.Drawing.Point(220, 10)
        $existingCategoryRadio.Size = New-Object System.Drawing.Size(120, 20)
        $existingCategoryRadio.Text = "既存のカテゴリ"
        $existingCategoryRadio.Checked = $true
        $editForm.Controls.Add($existingCategoryRadio)

        $newCategoryRadio = New-Object System.Windows.Forms.RadioButton
        $newCategoryRadio.Location = New-Object System.Drawing.Point(350, 10)
        $newCategoryRadio.Size = New-Object System.Drawing.Size(120, 20)
        $newCategoryRadio.Text = "新しいカテゴリ"
        $editForm.Controls.Add($newCategoryRadio)

        # 既存カテゴリ選択用コンボボックス
        $categoryComboBox = New-Object System.Windows.Forms.ComboBox
        $categoryComboBox.Location = New-Object System.Drawing.Point(220, 35)
        $categoryComboBox.Size = New-Object System.Drawing.Size(250, 20)
        $categoryComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $categories = Get-ChildItem -Path "data" -Directory | Select-Object -ExpandProperty Name
        $categoryComboBox.Items.AddRange($categories)
        $categoryComboBox.SelectedItem = $selectedPrompt.Category
        $editForm.Controls.Add($categoryComboBox)

        # 新しいカテゴリ入力用テキストボックス
        $newCategoryTextBox = New-Object System.Windows.Forms.TextBox
        $newCategoryTextBox.Location = New-Object System.Drawing.Point(220, 60)
        $newCategoryTextBox.Size = New-Object System.Drawing.Size(250, 20)
        $newCategoryTextBox.Enabled = $false
        $editForm.Controls.Add($newCategoryTextBox)

        # ラジオボタンの状態変更時の処理
        $existingCategoryRadio.Add_CheckedChanged({
            $categoryComboBox.Enabled = $existingCategoryRadio.Checked
            $newCategoryTextBox.Enabled = -not $existingCategoryRadio.Checked
        })

        $newCategoryRadio.Add_CheckedChanged({
            $categoryComboBox.Enabled = -not $newCategoryRadio.Checked
            $newCategoryTextBox.Enabled = $newCategoryRadio.Checked
        })

        # プロンプトラベル
        $promptLabel = New-Object System.Windows.Forms.Label
        $promptLabel.Location = New-Object System.Drawing.Point(10, 90)
        $promptLabel.Size = New-Object System.Drawing.Size(100, 20)
        $promptLabel.Text = "プロンプト"
        $editForm.Controls.Add($promptLabel)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = New-Object System.Drawing.Size(570, 300)  # テキストボックスのサイズを大きくしました
        $textBox.Location = New-Object System.Drawing.Point(10, 110)
        $textBox.Text = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($selectedPrompt.Content))
        $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                          [System.Windows.Forms.AnchorStyles]::Left -bor 
                          [System.Windows.Forms.AnchorStyles]::Right

        # テキストの選択を解除し、カーソルを先頭に移動
        $textBox.Select(0, 0)

        $saveButton = New-Object System.Windows.Forms.Button
        $saveButton.Size = New-Object System.Drawing.Size(100, 30)
        $saveButton.Text = "保存"
        $saveButton.Add_Click({
            $newCategory = if ($existingCategoryRadio.Checked) {
                $categoryComboBox.SelectedItem
            } else {
                $newCategoryTextBox.Text.Trim()
            }
    
            if ([string]::IsNullOrWhiteSpace($newCategory)) {
                [System.Windows.Forms.MessageBox]::Show("カテゴリを選択または入力してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
    
            $oldFilePath = $selectedPrompt.FilePath
            $oldCategory = $selectedPrompt.Category
            $categoryChanged = $oldCategory -ne $newCategory
    
            $selectedPrompt.Content = $textBox.Text
            $selectedPrompt.Category = $newCategory
    
            # カテゴリが変更された場合のみ、新しいファイルパスを設定
            if ($categoryChanged) {
                # カテゴリフォルダが存在しない場合は作成
                $newCategoryPath = Join-Path "data" $newCategory
                if (-not (Test-Path $newCategoryPath)) {
                    New-Item -ItemType Directory -Path $newCategoryPath | Out-Null
                }
    
                # 新しいファイル名とパスを設定
                $newFileName = Split-Path -Leaf $selectedPrompt.FileName
                $selectedPrompt.FileName = Join-Path $newCategory $newFileName
                $selectedPrompt.FilePath = Join-Path "data" $selectedPrompt.FileName
            }
    
            # プロンプトを保存
            Save-Prompt $selectedPrompt
            
            # カテゴリが変更された場合のみ、古いファイルを削除
            if ($categoryChanged) {
                Remove-Item -Path $oldFilePath -Force
            }
    
            [System.Windows.Forms.MessageBox]::Show(
                "プロンプトが保存されました。", 
                "保存完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $editForm.Close()
            # カテゴリリストとプロンプトリストを更新
            Update-CategoryList
            Update-PromptList
        })

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
        $cancelButton.Text = "キャンセル"
        $cancelButton.Add_Click({
            $editForm.Close()
        })

        # ボタンを配置する関数
        function Update-ButtonPositions {
            $buttonSpacing = 10
            $totalWidth = $saveButton.Width + $cancelButton.Width + $buttonSpacing
            $startX = ($editForm.ClientSize.Width - $totalWidth) / 2
            
            $saveButton.Left = $startX
            $saveButton.Top = $editForm.ClientSize.Height - $saveButton.Height - 10
            
            $cancelButton.Left = $startX + $saveButton.Width + $buttonSpacing
            $cancelButton.Top = $saveButton.Top
        }

        # 初期位置の設定
        Update-ButtonPositions

        # フォームのリサイズイベントを設定
        $editForm.Add_Resize({
            Update-ButtonPositions
        })

        $editForm.Controls.Add($textBox)
        $editForm.Controls.Add($saveButton)
        $editForm.Controls.Add($cancelButton)
        $editForm.ShowDialog()
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "プロンプトが選択されていません。", 
            "エラー", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Show-PromptPreview {
    <#
    .SYNOPSIS
        選択されたプロンプトのプレビューを表示します。
    .DESCRIPTION
        リストボックスで選択されたプロンプトの内容を
        読み取り専用のテキストボックスで表示します。
        また、LLMで確認するためのボタンをテキストボックスの横に配置します。
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = $script:FilteredPrompts[$selectedIndex]
        $previewForm = New-Object System.Windows.Forms.Form
        $previewForm.Text = "プロンプトプレビュー"
        $previewForm.Size = New-Object System.Drawing.Size(500, 420)
        $previewForm.StartPosition = "CenterScreen"

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = New-Object System.Drawing.Size(460, 330)
        $textBox.Location = New-Object System.Drawing.Point(10, 10)
        $textBox.Text = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($selectedPrompt.Content))
        $textBox.ReadOnly = $true
        $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                          [System.Windows.Forms.AnchorStyles]::Left -bor 
                          [System.Windows.Forms.AnchorStyles]::Right

        # テキストの選択を解除し、カーソルを先頭に移動
        $textBox.Select(0, 0)

        $previewForm.Controls.Add($textBox)

        # LLM送信ボタンの作成
        $llmSendButton = New-Object System.Windows.Forms.Button
        $llmSendButton.Size = New-Object System.Drawing.Size(100, 30)
        $llmSendButton.Text = "LLMで確認"
        $llmSendButton.Add_Click({ Show-LLMSendForm -PromptContent $textBox.Text })
        $previewForm.Controls.Add($llmSendButton)
        # ボタンを配置する関数
        function Update-ButtonPosition {
            $llmSendButton.Left = ($previewForm.ClientSize.Width - $llmSendButton.Width) / 2
            $llmSendButton.Top = $previewForm.ClientSize.Height - $llmSendButton.Height - 10
        }

        # 初期位置の設定
        Update-ButtonPosition

        # フォームのリサイズイベントを設定
        $previewForm.Add_Resize({
            Update-ButtonPosition
        })

        $previewForm.ShowDialog()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "プロンプトが選択されていません。", 
            "エラー", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Copy-SelectedPrompt {
    <#
    .SYNOPSIS
        選択されたプロンプトの内容をクリップボードにコピーします。
    .DESCRIPTION
        リストボックスで選択されたプロンプトの内容を
        クリップボードにコピーし、結果をメッセージボックスで表示します。
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = $script:FilteredPrompts[$selectedIndex]
        $result = Copy-TextToClipboard -Text $selectedPrompt.Content
        if ($result.Success) {
            [System.Windows.Forms.MessageBox]::Show(
                $result.Message, 
                "コピー完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                $result.Message, 
                "エラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "プロンプトが選択されていません。", 
            "エラー", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Update-PromptList {
    <#
    .SYNOPSIS
        プロンプトリストを更新します。
    .DESCRIPTION
        Get-Prompts 関数を使用して最新のプロンプト一覧を取得し、
        リストボックスの内容を更新します。フィルタリングされたプロンプトリストも保持します。
    #>
    $script:PromptListBox.Items.Clear()
    $allPrompts = Get-Prompts -Encoding UTF8
    $selectedCategory = $script:CategoryFilter.SelectedItem

    $script:FilteredPrompts = @()
    foreach ($prompt in $allPrompts) {
        if ($selectedCategory -eq "すべて" -or $prompt.Category -eq $selectedCategory) {
            $displayText = "$($prompt.Category): $($prompt.Title)"
            [void]$script:PromptListBox.Items.Add($displayText)
            $script:FilteredPrompts += $prompt
        }
    }
}

function Update-CategoryList {
    <#
    .SYNOPSIS
        カテゴリリストを更新します。
    .DESCRIPTION
        Get-Prompts 関数を使用して最新のカテゴリ一覧を取得し、
        カテゴリフィルタのコンボボックスを更新します。
    #>
    $prompts = Get-Prompts -Encoding UTF8
    $categories = @("すべて") + ($prompts | Select-Object -ExpandProperty Category -Unique)
    $script:CategoryFilter.Items.Clear()
    $script:CategoryFilter.Items.AddRange($categories)

    # コンボボックスのハンドルが作成された後に SelectedIndex を設定
    if ($script:CategoryFilter.IsHandleCreated) {
        $script:CategoryFilter.SelectedIndex = 0
    } else {
        # ハンドルが作成されていない場合は Loaded イベントで設定
        $script:CategoryFilter.Add_HandleCreated({
            $script:CategoryFilter.SelectedIndex = 0
        })
    }
}

function Remove-SelectedPrompt {
    <#
    .SYNOPSIS
        選択されたプロンプトを削除します。
    .DESCRIPTION
        リストボックスで選択されたプロンプトの削除確認ダイアログを表示し、
        ユーザーが確認した場合にプロンプトを削除します。
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = $script:FilteredPrompts[$selectedIndex]
        $result = [System.Windows.Forms.MessageBox]::Show(
            "本当に「$($selectedPrompt.Title)」を削除しますか？",
            "削除の確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $deleted = Remove-Prompt $selectedPrompt
            if ($deleted) {
                [System.Windows.Forms.MessageBox]::Show(
                    "プロンプトが削除されました。", 
                    "削除完了", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                # カテゴリリストを更新
                Update-CategoryList
                # リストボックスを更新
                Update-PromptList
                
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "プロンプトの削除に失敗しました。", 
                    "エラー", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "プロンプトが選択されていません。", 
            "エラー", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Exit-Application {
    <#
    .SYNOPSIS
        アプリケーションを終了します。
    .DESCRIPTION
        トレイアイコンを非表示にし、メインフォームを閉じ、
        アプリケーションを終了します。
    #>
    $script:TrayIcon.Visible = $false
    $script:TrayIcon.Dispose()
    if ($null -ne $script:MainForm -and -not $script:MainForm.IsDisposed) {
        $script:MainForm.Close()
    }
    [System.Windows.Forms.Application]::Exit()
}

function Show-CategoryManagementForm {
    <#
    .SYNOPSIS
        カテゴリ管理フォームを表示します。
    .DESCRIPTION
        カテゴリの一覧表示、編集、削除などの機能を提供するフォームを表示します。
    #>
    
    # カテゴリ管理フォームの作成
    $categoryForm = New-Object System.Windows.Forms.Form
    $categoryForm.Text = "カテゴリ管理"
    $categoryForm.Size = New-Object System.Drawing.Size(400,300)
    $categoryForm.StartPosition = "CenterScreen"

    # カテゴリ一覧を表示するListBoxの作成
    $categoryListBox = New-Object System.Windows.Forms.ListBox
    $categoryListBox.Location = New-Object System.Drawing.Point(10,10)
    $categoryListBox.Size = New-Object System.Drawing.Size(260,240)
    $categoryListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                              [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                              [System.Windows.Forms.AnchorStyles]::Left -bor 
                              [System.Windows.Forms.AnchorStyles]::Right
    $categoryForm.Controls.Add($categoryListBox)

    # 編集ボタンの作成
    $editButton = New-Object System.Windows.Forms.Button
    $editButton.Location = New-Object System.Drawing.Point(280,10)
    $editButton.Size = New-Object System.Drawing.Size(100,30)
    $editButton.Text = "編集"
    $editButton.Add_Click({
        $selectedCategory = $categoryListBox.SelectedItem
        if ($selectedCategory) {
            $newName = [Microsoft.VisualBasic.Interaction]::InputBox("新しいカテゴリ名を入力してください", "カテゴリ名の編集", $selectedCategory)
            if ($newName -and ($newName -ne $selectedCategory)) {
                $result = Rename-Category -OldName $selectedCategory -NewName $newName
                if ($result) {
                    [System.Windows.Forms.MessageBox]::Show("カテゴリ名を変更しました。", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    Update-CategoryListBox
                    Update-CategoryList
                    Update-PromptList
                } else {
                    [System.Windows.Forms.MessageBox]::Show("カテゴリ名の変更に失敗しました。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("カテゴリを選択してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $categoryForm.Controls.Add($editButton)

    # 削除ボタンの作成
    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Location = New-Object System.Drawing.Point(280,50)
    $deleteButton.Size = New-Object System.Drawing.Size(100,30)
    $deleteButton.Text = "削除"
    $deleteButton.Add_Click({
        $selectedCategory = $categoryListBox.SelectedItem
        if ($selectedCategory) {
            $result = Remove-Category -CategoryName $selectedCategory
            if ($result) {
                [System.Windows.Forms.MessageBox]::Show("カテゴリを削除しました。", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                Update-CategoryListBox
                Update-CategoryList
                Update-PromptList
            } else {
                [System.Windows.Forms.MessageBox]::Show("カテゴリの削除に失敗しました。フォルダが空でない可能性があります。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("カテゴリを選択してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $categoryForm.Controls.Add($deleteButton)

    # カテゴリ一覧を取得して表示する関数
    function Update-CategoryListBox {
        $categoryListBox.Items.Clear()
        $categories = Get-CategoryList
        foreach ($category in $categories) {
            $categoryListBox.Items.Add($category)
        }
    }

    # 初期表示時にカテゴリ一覧を更新
    Update-CategoryListBox

    # フォームを表示
    $categoryForm.ShowDialog()
}

function Get-CategoryList {
    <#
    .SYNOPSIS
        カテゴリの一覧を取得します。
    .DESCRIPTION
        dataフォルダ内のサブフォルダ名をカテゴリとして取得します。
    #>
    $categories = Get-ChildItem -Path "data" -Directory | Select-Object -ExpandProperty Name
    return $categories
}

function Rename-Category {
    <#
    .SYNOPSIS
        カテゴリ名を変更します。
    .DESCRIPTION
        指定された古いカテゴリ名を新しいカテゴリ名に変更し、
        関連するプロンプトファイルも移動します。
    .PARAMETER OldName
        変更前のカテゴリ名
    .PARAMETER NewName
        変更後のカテゴリ名
    .RETURNS
        成功した場合は $true、失敗した場合は $false を返します。
    #>
    param(
        [string]$OldName,
        [string]$NewName
    )

    try {
        $oldPath = Join-Path "data" $OldName
        $newPath = Join-Path "data" $NewName

        # カテゴリフォルダの名前を変更
        Rename-Item -Path $oldPath -NewName $NewName -ErrorAction Stop

        # プロンプトファイルのカテゴリ情報を更新
        $promptFiles = Get-ChildItem -Path $newPath -Filter "*.md"
        foreach ($file in $promptFiles) {
            $content = Get-Content $file.FullName -Raw -Encoding UTF8
            $updatedContent = $content -replace "Category:\s*$OldName", "Category: $NewName"
            Set-Content -Path $file.FullName -Value $updatedContent -NoNewline -Encoding UTF8
        }

        return $true
    }
    catch {
        Write-Error "カテゴリ名の変更中にエラーが発生しました: $_"
        return $false
    }
}

function Remove-Category {
    <#
    .SYNOPSIS
        空のカテゴリを削除します。
    .DESCRIPTION
        指定されたカテゴリが空の場合のみ、そのカテゴリを削除します。
    .PARAMETER CategoryName
        削除するカテゴリ名
    .RETURNS
        成功した場合は $true、失敗した場合は $false を返します。
    #>
    param(
        [string]$CategoryName
    )

    try {
        $categoryPath = Join-Path "data" $CategoryName

        # カテゴリフォルダが存在するか確認
        if (-not (Test-Path $categoryPath)) {
            Write-Error "指定されたカテゴリが存在しません: $CategoryName"
            return $false
        }

        # フォルダが空かどうか確認
        $files = Get-ChildItem -Path $categoryPath -File
        if ($files.Count -gt 0) {
            Write-Error "カテゴリフォルダが空ではありません: $CategoryName"
            return $false
        }

        # 空のフォルダを削除
        Remove-Item -Path $categoryPath -Force -Recurse
        return $true
    }
    catch {
        Write-Error "カテゴリの削除中にエラーが発生しました: $_"
        return $false
    }
}

function Show-ApiKeySettingsForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "API Key設定"
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"

    $openAILabel = New-Object System.Windows.Forms.Label
    $openAILabel.Location = New-Object System.Drawing.Point(10,20)
    $openAILabel.Size = New-Object System.Drawing.Size(80,20)
    $openAILabel.Text = "OpenAI:"
    $form.Controls.Add($openAILabel)

    $openAITextBox = New-Object System.Windows.Forms.TextBox
    $openAITextBox.Location = New-Object System.Drawing.Point(100,20)
    $openAITextBox.Size = New-Object System.Drawing.Size(180,20)
    $openAITextBox.Text = Get-ApiKey "OpenAI"
    $form.Controls.Add($openAITextBox)

    $claudeLabel = New-Object System.Windows.Forms.Label
    $claudeLabel.Location = New-Object System.Drawing.Point(10,50)
    $claudeLabel.Size = New-Object System.Drawing.Size(80,20)
    $claudeLabel.Text = "Claude:"
    $form.Controls.Add($claudeLabel)

    $claudeTextBox = New-Object System.Windows.Forms.TextBox
    $claudeTextBox.Location = New-Object System.Drawing.Point(100,50)
    $claudeTextBox.Size = New-Object System.Drawing.Size(180,20)
    $claudeTextBox.Text = Get-ApiKey "Claude"
    $form.Controls.Add($claudeTextBox)

    $geminiLabel = New-Object System.Windows.Forms.Label
    $geminiLabel.Location = New-Object System.Drawing.Point(10,80)
    $geminiLabel.Size = New-Object System.Drawing.Size(80,20)
    $geminiLabel.Text = "Gemini:"
    $form.Controls.Add($geminiLabel)

    $geminiTextBox = New-Object System.Windows.Forms.TextBox
    $geminiTextBox.Location = New-Object System.Drawing.Point(100,80)
    $geminiTextBox.Size = New-Object System.Drawing.Size(180,20)
    $geminiTextBox.Text = Get-ApiKey "Gemini"
    $form.Controls.Add($geminiTextBox)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(100,120)
    $saveButton.Size = New-Object System.Drawing.Size(75,23)
    $saveButton.Text = "保存"
    $saveButton.Add_Click({
        Set-ApiKey "OpenAI" $openAITextBox.Text
        Set-ApiKey "Claude" $claudeTextBox.Text
        Set-ApiKey "Gemini" $geminiTextBox.Text
        [System.Windows.Forms.MessageBox]::Show("APIキーが保存されました。")
        $form.Close()
    })
    $form.Controls.Add($saveButton)

    $form.ShowDialog()
}

function Show-LLMSendForm {
    param (
        [string]$PromptContent
    )

    $llmSendForm = New-Object System.Windows.Forms.Form
    $llmSendForm.Text = "LLM送信"
    $llmSendForm.Size = New-Object System.Drawing.Size(800, 600)
    $llmSendForm.StartPosition = "CenterScreen"
    $llmSendForm.AutoSize = $true
    $llmSendForm.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

    # プロンプト表示エリア
    $promptTextBox = New-Object System.Windows.Forms.TextBox
    $promptTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $promptTextBox.Size = New-Object System.Drawing.Size(380, 560)
    $promptTextBox.Multiline = $true
    $promptTextBox.ScrollBars = "Vertical"
    $promptTextBox.Text = $PromptContent
    $promptTextBox.ReadOnly = $false
    $llmSendForm.Controls.Add($promptTextBox)

    # 右側のパネル
    $rightPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $rightPanel.Location = New-Object System.Drawing.Point(400, 10)
    $rightPanel.Size = New-Object System.Drawing.Size(380, 540)
    $rightPanel.AutoSize = $true
    $rightPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $rightPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $rightPanel.WrapContents = $false
    $llmSendForm.Controls.Add($rightPanel)

    # LLM選択チェックボックスとモデル選択ドロップダウン
    $providers = @("OpenAI", "Claude", "Gemini")
    $checkBoxes = [ordered]@{}
    $modelDropDowns = [ordered]@{}

    foreach ($provider in $providers) {
        $providerPanel = New-Object System.Windows.Forms.Panel
        $providerPanel.Size = New-Object System.Drawing.Size(360, 30)
        $providerPanel.AutoSize = $true

        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Location = New-Object System.Drawing.Point(0, 0)
        $checkBox.Size = New-Object System.Drawing.Size(100, 20)
        $checkBox.Text = $provider
        $providerPanel.Controls.Add($checkBox)
        $checkBoxes[$provider] = $checkBox

        $modelDropDown = New-Object System.Windows.Forms.ComboBox
        $modelDropDown.Location = New-Object System.Drawing.Point(110, 0)
        $modelDropDown.Size = New-Object System.Drawing.Size(250, 20)
        $modelDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $providerPanel.Controls.Add($modelDropDown)
        $modelDropDowns[$provider] = $modelDropDown

        $models = Get-LLMModels -AIProvider $provider
        foreach ($model in $models) {
            $modelDropDown.Items.Add($model) | Out-Null
        }
        if ($modelDropDown.Items.Count -gt 0) {
            $modelDropDown.SelectedIndex = 0
        }

        $rightPanel.Controls.Add($providerPanel)
    }

    # 結果表示エリア（FlowLayoutPanel）
    $resultFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $resultFlowPanel.Location = New-Object System.Drawing.Point(0, ($yPos + 10))
    $resultFlowPanel.Size = New-Object System.Drawing.Size(360, 400)
    $resultFlowPanel.AutoSize = $true
    $resultFlowPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $resultFlowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $resultFlowPanel.WrapContents = $true
    $rightPanel.Controls.Add($resultFlowPanel)

    # チェックボックスの状態変更時のイベントハンドラ
    $providers | ForEach-Object {
        $provider = $_
        $checkBoxes[$provider].Add_CheckedChanged({
            UpdateResultPanel
        })
    }

    function UpdateResultPanel {
        $resultFlowPanel.Controls.Clear()
        foreach ($provider in $providers) {
            if ($checkBoxes[$provider].Checked) {
                $resultPanel = New-Object System.Windows.Forms.Panel
                $resultPanel.Size = New-Object System.Drawing.Size(360, 300)
                $resultPanel.AutoSize = $true
    
                $providerLabel = New-Object System.Windows.Forms.Label
                $providerLabel.Location = New-Object System.Drawing.Point(0, 0)
                $providerLabel.Size = New-Object System.Drawing.Size(340, 20)
                $providerLabel.Text = $provider
                $resultPanel.Controls.Add($providerLabel)
    
                $resultTextBox = New-Object System.Windows.Forms.TextBox
                $resultTextBox.Location = New-Object System.Drawing.Point(0, 25)
                $resultTextBox.Size = New-Object System.Drawing.Size(300, 400)
                $resultTextBox.Multiline = $true
                $resultTextBox.ScrollBars = "Vertical"
                $resultTextBox.ReadOnly = $true
                $resultTextBox.Tag = $provider
                $resultPanel.Controls.Add($resultTextBox)
    
                $resultFlowPanel.Controls.Add($resultPanel)
            }
        }
        $llmSendForm.PerformLayout()
        
        # フォームのサイズを再計算
        $newWidth = [Math]::Max(800, $rightPanel.Right + 20)
        $newHeight = [Math]::Max(600, $rightPanel.Bottom + 40)
        $llmSendForm.Size = New-Object System.Drawing.Size($newWidth, $newHeight)

        # フォームが現在表示されているスクリーンを取得
        $currentScreen = [System.Windows.Forms.Screen]::FromControl($llmSendForm)

        # フォームの位置を現在のスクリーンの中央に調整
        $workingArea = $currentScreen.WorkingArea
        $newX = $workingArea.Left + [int](($workingArea.Width - $newWidth) / 2)
        $newY = $workingArea.Top + [int](($workingArea.Height - $newHeight) / 2)
        
        $llmSendForm.Location = New-Object System.Drawing.Point($newX, $newY)
    }

    # 送信ボタン
    $sendButton = New-Object System.Windows.Forms.Button
    $sendButton.Size = New-Object System.Drawing.Size(100, 30)
    $sendButton.Text = "送信"
    $sendButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $rightPanel.Controls.Add($sendButton)

    $sendButton.Add_Click({
        $selectedProviders = $providers | Where-Object { $checkBoxes[$_].Checked }
        if ($selectedProviders.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("少なくとも1つのLLMを選択してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        foreach ($provider in $selectedProviders) {
            $resultTextBox = $resultFlowPanel.Controls | 
                Where-Object { $_ -is [System.Windows.Forms.Panel] } | 
                ForEach-Object { $_.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Tag -eq $provider } } | 
                Select-Object -First 1
    
            if ($resultTextBox) {
                $selectedModel = $modelDropDowns[$provider].SelectedItem
                $resultTextBox.Clear()
                $resultTextBox.AppendText("${provider} (${selectedModel})に送信中..`n")
                try {
                    $response = Invoke-AICompletion -AIProvider $provider -Prompt $promptTextBox.Text -Model $selectedModel
                    if ($response) {
                        $resultTextBox.Clear()
                        $resultTextBox.AppendText($response)
                    } else {
                        $resultTextBox.Clear()
                        $resultTextBox.AppendText("応答の取得に失敗しました。")
                    }
                }
                catch {
                    $resultTextBox.Clear()
                    $resultTextBox.AppendText("応答の取得中にエラーが発生しました: $_")
                }
            }
        }
    })

    $llmSendForm.ShowDialog()
}

Export-ModuleMember -Function Show-PromptManagerMainWindow, Show-PromptPreview, Edit-SelectedPrompt, Copy-SelectedPrompt, Show-SettingsForm, Show-NewPromptForm, Initialize-TrayIcon, Exit-Application, Show-CategoryManagementForm, Show-LLMSendForm