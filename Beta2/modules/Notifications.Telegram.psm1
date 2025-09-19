# modules/Notifications.Telegram.psm1
# Helper to send plain text messages to a Telegram chat

function Send-TelegramMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$ChatId,
        [Parameter(Mandatory)][string]$Text,
        [switch]$Silent
    )

    if ([string]::IsNullOrWhiteSpace($Token)) { throw 'Telegram bot token is empty.' }
    if ([string]::IsNullOrWhiteSpace($ChatId)) { throw 'Telegram chat id is empty.' }

    $uri = "https://api.telegram.org/bot{0}/sendMessage" -f $Token
    $maxLen = 3900

    $normalized = $Text -replace "`r`n", "`n"
    $normalized = $normalized.Trim()
    if ($normalized.Length -eq 0) { return }

    $chunks = @()
    $pending = $normalized
    while ($pending.Length -gt $maxLen) {
        $split = $pending.LastIndexOf("`n", $maxLen)
        if ($split -lt 0) { $split = $maxLen }
        $chunk = $pending.Substring(0, $split).Trim()
        if ($chunk.Length -gt 0) { $chunks += $chunk }
        $pending = $pending.Substring($split)
        $pending = $pending.TrimStart("`n")
    }
    if ($pending.Length -gt 0) { $chunks += $pending }

    foreach ($chunk in $chunks) {
        $body = @{
            chat_id = $ChatId
            text    = $chunk
        }
        if ($Silent.IsPresent) { $body.disable_notification = $true }
        try {
            Invoke-RestMethod -Uri $uri -Method Post -Body $body -ErrorAction Stop | Out-Null
        }
        catch {
            $msg = $_.Exception.Message
            throw "Telegram send failed: $msg"
        }
    }
}

Export-ModuleMember -Function Send-TelegramMessage
