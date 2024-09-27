<#
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
        $script:MainForm.Size = New-Object System.Drawing.Size(500,300)
        $script:MainForm.StartPosition = "CenterScreen"

        # リストボックスの作成
        $script:PromptListBox = New-Object System.Windows.Forms.ListBox
        $script:PromptListBox.Location = New-Object System.Drawing.Point(10,10)
        $script:PromptListBox.Size = New-Object System.Drawing.Size(360,240)
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
        $script:PromptListBox.Size = New-Object System.Drawing.Size(360, 210)

        # プロンプトリストの更新
        Update-PromptList

        # 終了ボタンの追加
        $exitButton = New-Object System.Windows.Forms.Button
        $exitButton.Location = New-Object System.Drawing.Point(380, 220)
        $exitButton.Size = New-Object System.Drawing.Size(100, 30)
        $exitButton.Text = "終了"
        $exitButton.Add_Click({ 
            if ($null -ne $script:MainForm) {
                $script:MainForm.Hide()
            }
        })
        $exitButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor 
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
        if ($savedPrompt) {
            [System.Windows.Forms.MessageBox]::Show(
                "新しいプロンプトを保存しました。", 
                "保存完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # メインウィンドウのリストボックスを更新
            Update-PromptList
            
            $newPromptForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "プロンプトの保存に失敗しました。", 
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
        内容を編集して保存することができます。
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = (Get-Prompts -Encoding UTF8)[$selectedIndex]
        $editForm = New-Object System.Windows.Forms.Form
        $editForm.Text = "プロンプト編集"
        $editForm.Size = New-Object System.Drawing.Size(400, 350)
        $editForm.StartPosition = "CenterScreen"

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = New-Object System.Drawing.Size(380, 260)
        $textBox.Location = New-Object System.Drawing.Point(10, 10)
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
            $selectedPrompt.Content = $textBox.Text
            Save-Prompt $selectedPrompt
            [System.Windows.Forms.MessageBox]::Show(
                "プロンプトが保存されました。", 
                "保存完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $editForm.Close()
            # リストボックスを更新
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
            <#
            .SYNOPSIS
                保存ボタンとキャンセルボタンをフォームの中央下に配置します。
            .DESCRIPTION
                フォームのサイズ変更イベントなどで呼び出され、ボタンの位置を更新します。
            #>
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
    #>
    $selectedIndex = $script:PromptListBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $selectedPrompt = (Get-Prompts -Encoding UTF8)[$selectedIndex]
        $previewForm = New-Object System.Windows.Forms.Form
        $previewForm.Text = "プロンプトプレビュー"
        $previewForm.Size = New-Object System.Drawing.Size(400,300)
        $previewForm.StartPosition = "CenterScreen"

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = New-Object System.Drawing.Size(380,260)
        $textBox.Location = New-Object System.Drawing.Point(10,10)
        $textBox.Text = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($selectedPrompt.Content))
        $textBox.ReadOnly = $true
        $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                          [System.Windows.Forms.AnchorStyles]::Left -bor 
                          [System.Windows.Forms.AnchorStyles]::Right

        # テキストの選択を解除し、カーソルを先頭に移動
        $textBox.Select(0, 0)

        $previewForm.Controls.Add($textBox)
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
        $selectedPrompt = (Get-Prompts -Encoding UTF8)[$selectedIndex]
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
        リストボックスの内容を更新します。
    #>
    $script:PromptListBox.Items.Clear()
    $prompts = Get-Prompts -Encoding UTF8
    $selectedCategory = $script:CategoryFilter.SelectedItem

    foreach ($prompt in $prompts) {
        if ($selectedCategory -eq "すべて" -or $prompt.Category -eq $selectedCategory) {
            $displayText = "$($prompt.Category): $($prompt.Title)"
            [void]$script:PromptListBox.Items.Add($displayText)
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
        $selectedPrompt = (Get-Prompts -Encoding UTF8)[$selectedIndex]
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

Export-ModuleMember -Function Show-PromptManagerMainWindow, Show-PromptPreview, Edit-SelectedPrompt, Copy-SelectedPrompt, Show-SettingsForm, Show-NewPromptForm, Initialize-TrayIcon, Exit-Application