# Run-Backup.ps1
#requires -Version 5.1
chcp 65001 > $null

Write-Host "[INFO] Запуск процесса резервного копирования..." -ForegroundColor Yellow

$pipelinePath = Join-Path $PSScriptRoot 'core\Pipeline.psm1'
if (!(Test-Path $pipelinePath)) {
    Write-Error "Не найден Pipeline.psm1: $pipelinePath"
    exit 1
}
Import-Module -Force -DisableNameChecking $pipelinePath -ErrorAction Stop

$configRoot = Join-Path $PSScriptRoot 'config'
$basesDir   = Join-Path $configRoot 'bases'
$logDir     = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$bases = Get-ChildItem -Path $basesDir -Filter *.json -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName }
if (-not $bases -or $bases.Count -eq 0) {
    Write-Error "Не найдено ни одной базы в $basesDir"
    exit 1
}

$sessionLog = Join-Path $logDir ("backup_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] $Message"
    $line | Out-File -FilePath $sessionLog -Append -Encoding UTF8
    Write-Host $line
}

function ConvertTo-Hashtable {
    param([Parameter(Mandatory = $true)]$Object)
    if ($null -eq $Object) { return @{} }
    if ($Object -is [hashtable]) { return $Object }
    $ht = @{}
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) { $ht[$key] = $Object[$key] }
    }
    elseif ($Object.PSObject) {
        foreach ($p in $Object.PSObject.Properties) { $ht[$p.Name] = $p.Value }
    }
    return $ht
}

$startedAt = Get-Date

$settingsFile = Join-Path $configRoot 'settings.json'
$settingsData = @{}
if (Test-Path $settingsFile) {
    try {
        $settingsData = ConvertTo-Hashtable (Get-Content $settingsFile -Raw | ConvertFrom-Json)
        if (-not ($settingsData -is [hashtable])) { $settingsData = @{} }
    }
    catch {
        Write-Log ("[WARN] Не удалось прочитать settings.json: {0}" -f $_.Exception.Message)
        $settingsData = @{}
    }
}

$after = 3
if ($settingsData.ContainsKey('AfterBackup')) {
    try { $after = [int]$settingsData['AfterBackup'] } catch { $after = 3 }
}

$telegramEnabled = $false
$telegramChatId = ''
$telegramOnlyErrors = $false
if ($settingsData.ContainsKey('Telegram')) {
    $telegramRaw = ConvertTo-Hashtable $settingsData['Telegram']
    if ($telegramRaw.ContainsKey('Enabled')) { $telegramEnabled = [bool]$telegramRaw['Enabled'] }
    if ($telegramRaw.ContainsKey('ChatId')) { $telegramChatId = ('' + $telegramRaw['ChatId']).Trim() }
    if ($telegramRaw.ContainsKey('NotifyOnlyOnErrors')) { $telegramOnlyErrors = [bool]$telegramRaw['NotifyOnlyOnErrors'] }
}

$telegramToken = $null
$telegramModulePath = Join-Path $PSScriptRoot 'modules\Notifications.Telegram.psm1'
$cryptoModulePath   = Join-Path $PSScriptRoot 'modules\Common.Crypto.psm1'
if ($telegramEnabled) {
    if (-not (Test-Path $telegramModulePath)) {
        Write-Log "[WARN] Включены уведомления Telegram, но модуль Notifications.Telegram.psm1 не найден. Уведомления отключены."
        $telegramEnabled = $false
    }
    else {
        try {
            Import-Module -Force -DisableNameChecking $telegramModulePath -ErrorAction Stop
        }
        catch {
            Write-Log ("[WARN] Не удалось загрузить модуль Telegram: {0}" -f $_.Exception.Message)
            $telegramEnabled = $false
        }
    }

    if ($telegramEnabled) {
        if (-not (Test-Path $cryptoModulePath)) {
            Write-Log "[WARN] Не найден модуль Common.Crypto.psm1, уведомления Telegram отключены."
            $telegramEnabled = $false
        }
        else {
            try {
                Import-Module -Force -DisableNameChecking $cryptoModulePath -ErrorAction Stop
            }
            catch {
                Write-Log ("[WARN] Не удалось загрузить модуль Common.Crypto: {0}" -f $_.Exception.Message)
                $telegramEnabled = $false
            }
        }
    }

    if ($telegramEnabled) {
        $keyPath = Join-Path $configRoot 'key.bin'
        $secretsPath = Join-Path $configRoot 'secrets.json.enc'
        if ((Test-Path $keyPath) -and (Test-Path $secretsPath)) {
            try {
                $secrets = ConvertTo-Hashtable (Decrypt-Secrets -InFile $secretsPath -KeyPath $keyPath)
                if ($secrets.ContainsKey('__TELEGRAM_BOT_TOKEN')) {
                    $telegramToken = [string]$secrets['__TELEGRAM_BOT_TOKEN']
                }
            }
            catch {
                Write-Log ("[WARN] Не удалось расшифровать секреты для Telegram: {0}" -f $_.Exception.Message)
                $telegramEnabled = $false
            }
        }
        else {
            Write-Log "[WARN] Не найдены ключ или файл секретов. Уведомления Telegram отключены."
            $telegramEnabled = $false
        }
    }

    if ($telegramEnabled -and [string]::IsNullOrWhiteSpace($telegramChatId)) {
        Write-Log "[WARN] Не задан ChatId для Telegram. Уведомления отключены."
        $telegramEnabled = $false
    }
    if ($telegramEnabled -and [string]::IsNullOrWhiteSpace($telegramToken)) {
        Write-Log "[WARN] Не найден токен бота Telegram. Уведомления отключены."
        $telegramEnabled = $false
    }
}

$results = @()

foreach ($tag in $bases) {
    $entry = [ordered]@{
        Tag      = $tag
        Status   = 'Unknown'
        Artifact = $null
        Error    = $null
    }
    try {
        $ctx = @{
            Tag        = $tag
            ConfigRoot = $configRoot
            ConfigDir  = $basesDir
            Log        = { param($msg) Write-Log ("[{0}] {1}" -f $tag, $msg) }
        }
        $artifact = Invoke-Pipeline -Ctx $ctx
        if ($null -eq $artifact) {
            $entry.Status = 'Skipped'
        }
        else {
            $entry.Status = 'Success'
            $entry.Artifact = $artifact
        }
    }
    catch {
        $entry.Status = 'Error'
        $entry.Error = $_.Exception.Message
        Write-Log ("[ОШИБКА][{0}] {1}" -f $tag, $_.Exception.Message)
    }
    $results += [pscustomobject]$entry
}

$total       = $results.Count
$successCnt  = ($results | Where-Object { $_.Status -eq 'Success' }).Count
$skippedCnt  = ($results | Where-Object { $_.Status -eq 'Skipped' }).Count
$errorCnt    = ($results | Where-Object { $_.Status -eq 'Error' }).Count
$elapsed     = (Get-Date) - $startedAt
$elapsedText = "{0:hh\:mm\:ss}" -f $elapsed

Write-Log ("[INFO] Итог: всего {0}, успешно {1}, ошибок {2}, пропущено {3}" -f $total, $successCnt, $errorCnt, $skippedCnt)
Write-Log ("[INFO] Лог сохранён: {0}" -f $sessionLog)

if ($telegramEnabled) {
    $sendReport = $true
    if ($telegramOnlyErrors -and $errorCnt -eq 0) { $sendReport = $false }

    if ($sendReport -and (Get-Command Send-TelegramMessage -ErrorAction SilentlyContinue)) {
        $lines = @()
        $lines += "Отчёт о резервном копировании"
        $lines += ("Дата: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'))
        $lines += ("Баз всего: {0}, успешно: {1}, ошибок: {2}, пропущено: {3}" -f $total, $successCnt, $errorCnt, $skippedCnt)
        $lines += ("Длительность: {0}" -f $elapsedText)
        $lines += ("Лог: {0}" -f (Split-Path $sessionLog -Leaf))

        foreach ($item in $results) {
            $tag = $item.Tag
            switch ($item.Status) {
                'Success' {
                    $name = if ($item.Artifact) { Split-Path $item.Artifact -Leaf } else { '—' }
                    $lines += ("- {0}: успешно ({1})" -f $tag, $name)
                }
                'Skipped' {
                    $lines += ("- {0}: пропущена" -f $tag)
                }
                'Error' {
                    $err = $item.Error
                    if ($err.Length -gt 160) { $err = $err.Substring(0, 160).Trim() + '...' }
                    $lines += ("- {0}: ошибка ({1})" -f $tag, $err)
                }
                default {
                    $lines += ("- {0}: статус {1}" -f $tag, $item.Status)
                }
            }
        }

        $message = ($lines -join "`n")
        try {
            Send-TelegramMessage -Token $telegramToken -ChatId $telegramChatId -Text $message | Out-Null
            Write-Log "[INFO] Отчёт отправлен в Telegram."
        }
        catch {
            Write-Log ("[WARN] Не удалось отправить отчёт в Telegram: {0}" -f $_.Exception.Message)
        }
    }
    elseif (-not $sendReport) {
        Write-Log "[INFO] Отправка отчёта в Telegram пропущена: ошибок нет, включён режим 'только при ошибках'."
    }
    else {
        Write-Log "[WARN] Функция Send-TelegramMessage недоступна, отчёт не отправлен."
    }
}

switch ($after) {
    1 { Write-Log "[INFO] Выключаем ПК";  Stop-Computer -Force }
    2 { Write-Log "[INFO] Перезагружаем ПК"; Restart-Computer -Force }
    default { Write-Log "[INFO] Завершено. Действий с ПК нет." }
}

