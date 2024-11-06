# 設定ファイルのパスを取得
$settingsPath = Join-Path $PSScriptRoot "..\..\config\settings.xml"

# 設定ファイルの読み込みとエラーハンドリング
try {
    $settings = [xml](Get-Content $settingsPath)
    
    # 必須の設定ノードの存在確認
    $llmSettingsNode = $settings.SelectSingleNode("/Settings/LLMSettings")
    if (-not $llmSettingsNode) {
        throw "LLM設定が見つかりません"
    }
} catch {
    throw "設定ファイルの読み込みに失敗しました: $settingsPath`nエラー: $_"
}

function Get-LLMModels {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("OpenAI", "Claude", "Gemini")]
        [string]$AIProvider
    )

    $llmSettings = $settings.Settings.LLMSettings
    return $llmSettings.$AIProvider.Models.Model
}

function ConvertTo-UTF8 {
    param([string]$inputString)
    return [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($inputString))
}

function Invoke-AICompletion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("OpenAI", "Claude", "Gemini")]
        [string]$AIProvider,

        [Parameter(Mandatory=$true)]
        [string]$Prompt,

        [Parameter(Mandatory=$true)]
        [string]$Model
    )

    $headers = @{
        "Content-Type" = "application/json; charset=utf-8"
    }

    $llmSettings = $settings.Settings.LLMSettings

    switch ($AIProvider) {
        "OpenAI" {
            $endpoint = $llmSettings.OpenAI.Endpoint
            $headers["Authorization"] = "Bearer $(Get-ApiKey -Provider 'OpenAI')"
            $body = @{
                model = $Model
                messages = @(
                    @{
                        role = "user"
                        content = $Prompt
                    }
                )
            }
        }
        "Claude" {
            $endpoint = $llmSettings.Claude.Endpoint
            $headers["anthropic-version"] = "2023-06-01"
            $headers["x-api-key"] = Get-ApiKey -Provider 'Claude'
            $body = @{
                model = $Model
                messages = @(
                    @{
                        role = "user"
                        content = $Prompt
                    }
                )
                max_tokens = 1000
            }
        }
        "Gemini" {
            $endpoint = $llmSettings.Gemini.Endpoint -f $Model
            $endpoint += "?key=$(Get-ApiKey -Provider 'Gemini')"
            $headers["Content-Type"] = "application/json; charset=utf-8"
            $body = @{
                contents = @(
                    @{
                        parts = @(
                            @{
                                text = $Prompt
                            }
                        )
                    }
                )
            }
        }
    }

    try {
        $response = Invoke-APIRequest -Endpoint $endpoint -Headers $headers -Body $body
        # 応答の処理
        switch ($AIProvider) {
            "OpenAI" { 
                $content = $response.choices[0].message.content
                return ConvertTo-UTF8 $content
            }
            "Claude" { 
                $content = $response.content[0].text
                return ConvertTo-UTF8 $content
            }
            "Gemini" { return $response.candidates[0].content.parts[0].text }
        }
    }
    catch {
        Write-Error "API呼び出し中にエラーが発生しました: $_"
        return $null
    }
}

function Invoke-APIRequest {
    param(
        [string]$Endpoint,
        [hashtable]$Headers,
        [object]$Body,
        [int]$Timeout = 30
    )

    try {
        $jsonBody = $Body | ConvertTo-Json -Depth 10
        $jsonBodyBytes = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
        
        $response = Invoke-RestMethod `
            -Uri $Endpoint `
            -Method Post `
            -Headers $Headers `
            -Body $jsonBodyBytes `
            -ContentType "application/json; charset=utf-8" `
            -TimeoutSec $Timeout `
            -ErrorAction Stop

        return $response
    }
    catch {
        Write-Error "API呼び出し中にエラーが発生しました: $($_.Exception.Message)"
        throw
    }
}

# すべての必要な関数をエクスポート
Export-ModuleMember -Function Get-LLMModels, ConvertTo-UTF8, Invoke-AICompletion, Invoke-APIRequest
