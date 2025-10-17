# Run-Backup.ps1
#requires -Version 5.1
param(
    [Parameter(Mandatory = $false)]
    [string]$BaseName = ""
)

chcp 65001 > $null

if ($BaseName) {
    Write-Host "[INFO] Запуск резервного копирования для базы '$BaseName'..." -ForegroundColor Yellow
} else {
    Write-Host "[INFO] Запуск процесса резервного копирования..." -ForegroundColor Yellow
}

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

# Очистка старых логов (оставляем только последние 30)
$maxLogs = 30
$existingLogs = Get-ChildItem -Path $logDir -Filter "backup_*.log" -File | Sort-Object CreationTime -Descending
if ($existingLogs.Count -gt $maxLogs) {
    $toDelete = $existingLogs[$maxLogs..($existingLogs.Count - 1)]
    foreach ($oldLog in $toDelete) {
        Remove-Item $oldLog.FullName -Force -ErrorAction SilentlyContinue
    }
}

if ($BaseName) {
    # Run backup for specific database
    $baseConfigPath = Join-Path $basesDir "$BaseName.json"
    if (-not (Test-Path $baseConfigPath)) {
        Write-Error "База '$BaseName' не найдена в $basesDir"
        exit 1
    }
    $bases = @($BaseName)
    Write-Host "[INFO] Резервное копирование для базы: $BaseName" -ForegroundColor Green
} else {
    # Run backup for all databases
    $bases = Get-ChildItem -Path $basesDir -Filter *.json -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName }
    if (-not $bases -or $bases.Count -eq 0) {
        Write-Error "Не найдено ни одной базы в $basesDir"
        exit 1
    }
    Write-Host "[INFO] Найдено баз для резервного копирования: $($bases.Count)" -ForegroundColor Green
}

if ($BaseName) {
    $sessionLog = Join-Path $logDir ("backup_{0}_{1}.log" -f $BaseName, (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
} else {
    $sessionLog = Join-Path $logDir ("backup_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
}

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
$telegramSecondaryChatId = ''
$telegramOnlyErrors = $false
if ($settingsData.ContainsKey('Telegram')) {
    $telegramRaw = ConvertTo-Hashtable $settingsData['Telegram']
    if ($telegramRaw.ContainsKey('Enabled')) { $telegramEnabled = [bool]$telegramRaw['Enabled'] }
    if ($telegramRaw.ContainsKey('ChatId')) { $telegramChatId = ('' + $telegramRaw['ChatId']).Trim() }
    if ($telegramRaw.ContainsKey('SecondaryChatId')) { $telegramSecondaryChatId = ('' + $telegramRaw['SecondaryChatId']).Trim() }
    if ($telegramRaw.ContainsKey('NotifyOnlyOnErrors')) { $telegramOnlyErrors = [bool]$telegramRaw['NotifyOnlyOnErrors'] }
}

$telegramToken = $null
$telegramSecondaryToken = $null
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
                if ($secrets.ContainsKey('__TELEGRAM_SECONDARY_BOT_TOKEN')) {
                    $telegramSecondaryToken = [string]$secrets['__TELEGRAM_SECONDARY_BOT_TOKEN']
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

if ($telegramEnabled) {
    if (-not $telegramSecondaryToken -or [string]::IsNullOrWhiteSpace($telegramSecondaryChatId)) {
        $telegramSecondaryToken = $null
        $telegramSecondaryChatId = ''
    }
}
else {
    $telegramSecondaryToken = $null
    $telegramSecondaryChatId = ''
}

$results = @()
$hasErrors = $false

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
        $hasErrors = $true
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
Write-Host ("[INFO] Лог сохранён: {0}" -f $sessionLog)

if ($telegramEnabled) {
    $sendReport = $true
    if ($telegramOnlyErrors -and $errorCnt -eq 0) { $sendReport = $false }

    if ($sendReport -and (Get-Command Send-TelegramMessage -ErrorAction SilentlyContinue)) {
        try {
            $logLines = Get-Content -Path $sessionLog -Encoding UTF8
            $message  = ($logLines -join "`n").Trim()
            if ($message) {
                Send-TelegramMessage -Token $telegramToken -ChatId $telegramChatId -Text $message | Out-Null
                Write-Host "[INFO] Отчёт отправлен в Telegram."
                if ($telegramSecondaryToken -and $telegramSecondaryChatId) {
                    try {
                        Send-TelegramMessage -Token $telegramSecondaryToken -ChatId $telegramSecondaryChatId -Text $message | Out-Null
                        Write-Host "[INFO] Дополнительный отчёт отправлен ответственному."
                    }
                    catch {
                        Write-Host ("[WARN] Не удалось отправить отчёт ответственному: {0}" -f $_.Exception.Message)
                    }
                }
            }
            else {
                Write-Host "[WARN] Отчёт в Telegram не отправлен: лог пустой."
            }
        }
        catch {
            Write-Host ("[WARN] Не удалось отправить отчёт в Telegram: {0}" -f $_.Exception.Message)
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

# Возвращаем код ошибки для CMD
if ($hasErrors) {
    exit 1
} else {
    exit 0
}