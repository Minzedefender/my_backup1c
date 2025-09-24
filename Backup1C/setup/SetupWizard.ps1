#requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms
Import-Module -Force -DisableNameChecking (Join-Path $PSScriptRoot '..\modules\Common.Crypto.psm1')

# ---------- helpers ----------
function Select-FolderDialog($description) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $description
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath } else { throw 'Отменено' }
}
function Select-FileDialog($filter, $title, $initialDir = $null) {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = $filter
    $dlg.Title  = $title
    if ($initialDir -and (Test-Path $initialDir)) { $dlg.InitialDirectory = $initialDir }
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.FileName } else { throw 'Отменено' }
}
function Read-Choice($prompt, $choices) {
    Write-Host $prompt
    for ($i=0; $i -lt $choices.Count; $i++){ Write-Host ("{0} - {1}" -f ($i+1), $choices[$i]) }
    do { $x = Read-Host ("Ваш выбор (1/{0})" -f $choices.Count) }
    while (-not ($x -match '^\d+$') -or [int]$x -lt 1 -or [int]$x -gt $choices.Count)
    [int]$x
}
# Преобразуем PSCustomObject -> Hashtable (на случай, если дешифратор вернул не Hashtable)
function ConvertTo-Hashtable($obj){
    if ($obj -is [hashtable]) { return $obj }
    $ht = @{}
    if ($obj -is [System.Collections.IDictionary]) {
        foreach($k in $obj.Keys){ $ht[$k] = $obj[$k] }
    } elseif ($obj -and $obj.PSObject) {
        foreach($p in $obj.PSObject.Properties){ $ht[$p.Name] = $p.Value }
    }
    return $ht
}

# Только веб-службы (Apache/IIS)
function Get-WebServices {
    $masks = @('Apache2.4','Apache*','httpd*','W3SVC','WAS')  # IIS: W3SVC (WWW), WAS
    $all = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $n = $_.Name
        $masks | Where-Object { $n -like $_ } | Select-Object -First 1
    }
    $all | Sort-Object @{Expression='Status';Descending=$true},
                     @{Expression='DisplayName';Descending=$false}
}

# Поиск 1cestart.exe (x64/x86)
function Get-1CEStartCandidates {
    $list = @()
    $pf   = $env:ProgramFiles
    $pf86 = ${env:ProgramFiles(x86)}
    if ($pf)   { $list += (Join-Path $pf   '1cv8\common\1cestart.exe') }
    if ($pf86) { $list += (Join-Path $pf86 '1cv8\common\1cestart.exe') }
    $list | Where-Object { $_ -and (Test-Path $_) }
}
function Select-1CEStart {
    $cands = Get-1CEStartCandidates
    $init  = if ($cands -and $cands[0]) { Split-Path $cands[0] -Parent } else { "$env:ProgramFiles\1cv8\common" }
    Select-FileDialog "1cestart.exe|1cestart.exe" "Выберите 1cestart.exe (обычно C:\Program Files\1cv8\common)" $init
}

# ---------- paths ----------
$configRoot   = Join-Path $PSScriptRoot '..\config'
$basesDir     = Join-Path $configRoot  'bases'
$settingsFile = Join-Path $configRoot  'settings.json'
$keyPath      = Join-Path $configRoot  'key.bin'
$secretsFile  = Join-Path $configRoot  'secrets.json.enc'

if (-not (Test-Path $configRoot)) { New-Item -ItemType Directory -Path $configRoot | Out-Null }
if (-not (Test-Path $basesDir  )) { New-Item -ItemType Directory -Path $basesDir   | Out-Null }

# ---------- load & merge secrets ----------
[hashtable]$allSecrets = @{}
if ((Test-Path $secretsFile) -and (Test-Path $keyPath)) {
    try {
        $raw = Decrypt-Secrets -InFile $secretsFile -KeyPath $keyPath
        $allSecrets = ConvertTo-Hashtable $raw
        if (-not ($allSecrets -is [hashtable])) { $allSecrets = @{} }
    } catch { $allSecrets = @{} }
}

$telegramTokenKey = '__TELEGRAM_BOT_TOKEN'
$secondaryTokenKey = '__TELEGRAM_SECONDARY_BOT_TOKEN'
$settingsData = @{}
if (Test-Path $settingsFile) {
    try {
        $settingsData = ConvertTo-Hashtable (Get-Content $settingsFile -Raw | ConvertFrom-Json)
        if (-not ($settingsData -is [hashtable])) { $settingsData = @{} }
    } catch { $settingsData = @{} }
}

$telegramData = @{}
if ($settingsData.ContainsKey('Telegram')) {
    $telegramData = ConvertTo-Hashtable $settingsData['Telegram']
    if (-not ($telegramData -is [hashtable])) { $telegramData = @{} }
}

$telegramEnabledDefault = $false
if ($telegramData.ContainsKey('Enabled')) { $telegramEnabledDefault = [bool]$telegramData['Enabled'] }

$telegramChatIdDefault = ''
if ($telegramData.ContainsKey('ChatId')) { $telegramChatIdDefault = [string]$telegramData['ChatId'] }
if ($telegramChatIdDefault) { $telegramChatIdDefault = $telegramChatIdDefault.Trim() }

$telegramSecondaryChatDefault = ''
if ($telegramData.ContainsKey('SecondaryChatId')) { $telegramSecondaryChatDefault = [string]$telegramData['SecondaryChatId'] }
if ($telegramSecondaryChatDefault) { $telegramSecondaryChatDefault = $telegramSecondaryChatDefault.Trim() }

$telegramOnlyErrorsDefault = $false
if ($telegramData.ContainsKey('NotifyOnlyOnErrors')) { $telegramOnlyErrorsDefault = [bool]$telegramData['NotifyOnlyOnErrors'] }

$existingTelegramToken = ''
if ($allSecrets.ContainsKey($telegramTokenKey)) { $existingTelegramToken = [string]$allSecrets[$telegramTokenKey] }
$existingTelegramSecondaryToken = ''
if ($allSecrets.ContainsKey($secondaryTokenKey)) { $existingTelegramSecondaryToken = [string]$allSecrets[$secondaryTokenKey] }

$telegramOptions = [ordered]@{
    Enabled = $telegramEnabledDefault
    ChatId = $telegramChatIdDefault
    NotifyOnlyOnErrors = $telegramOnlyErrorsDefault
    SecondaryChatId = $telegramSecondaryChatDefault
}

$allConfigs = @()

# ---------- wizard ----------
while ($true) {
    $cfg = @{}
    $cfg.Tag = Read-Host "Введите уникальное имя базы (например, ShopDB)"

    $t = Read-Choice "Выберите тип бэкапа:" @('Копия файла .1CD','Выгрузка .dt через конфигуратор')
    $cfg.BackupType = if ($t -eq 1) { '1CD' } else { 'DT' }

    if ($cfg.BackupType -eq '1CD') {
        $cfg.SourcePath = Select-FileDialog "Файл 1Cv8.1CD|1Cv8.1CD" "Выберите файл 1Cv8.1CD (где ХРАНИТСЯ база)"
    } else {
        $cfg.SourcePath = Select-FolderDialog "Выберите каталог, в котором лежит база 1С"

        # 1cestart.exe (приоритет)
        try { $cfg.ExePath = Select-1CEStart } catch { $cfg.ExePath = $null }

        # Выбор веб-служб: Авто / Ручной / Нет
        $stopMode = Read-Choice "Останавливать веб-службы при выгрузке .dt?" @(
            'Авто (Apache2.4; при отсутствии — все найденные веб-службы)',
            'Выбрать вручную',
            'Нет'
        )
        switch ($stopMode) {
            1 {
                $web = Get-WebServices
                if (-not $web) {
                    Write-Host "Веб-службы не найдены" -ForegroundColor Yellow
                }
                elseif ($web.Name -contains 'Apache2.4') {
                    $cfg.StopServices = @('Apache2.4')
                    Write-Host "Будет остановлена служба: Apache2.4" -ForegroundColor DarkCyan
                }
                else {
                    $cfg.StopServices = $web | Select-Object -ExpandProperty Name
                    if ($cfg.StopServices) {
                        Write-Host ("Будут остановлены: {0}" -f ($cfg.StopServices -join ', ')) -ForegroundColor DarkCyan
                    }
                }
            }
            2 {
                $web = Get-WebServices
                if (-not $web) {
                    Write-Host "Веб-службы не найдены" -ForegroundColor Yellow
                }
                else {
                    Write-Host "`nДоступные веб-службы:" -ForegroundColor Cyan
                    $i = 1
                    foreach ($s in $web) {
                        Write-Host ("{0}) {1} [{2}] — {3}" -f $i, $s.DisplayName, $s.Name, $s.Status)
                        $i++
                    }
                    $raw = Read-Host "Введите номера через запятую (например, 1,3)"
                    $idx = $raw -split '[,; ]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object {[int]$_}
                    $sel = @()
                    for ($j=0; $j -lt $web.Count; $j++) { if ($idx -contains ($j+1)) { $sel += $web[$j].Name } }
                    if ($sel.Count -gt 0) {
                        $cfg.StopServices = $sel
                        Write-Host ("Будут остановлены: {0}" -f ($sel -join ', ')) -ForegroundColor DarkCyan
                    } else {
                        Write-Host "Ничего не выбрано — службы останавливаться не будут" -ForegroundColor Yellow
                    }
                }
            }
            default { $cfg.StopServices = @() }  # "Нет"
        }

        # Логин/пароль для .DT (не затираем чужие базы)
        $lg = Read-Host "Логин 1С (если не требуется — пусто)"
        $pw = Read-Host "Пароль 1С (если не требуется — пусто)"
        if ($lg) { $allSecrets["$($cfg.Tag)__DT_Login"]    = $lg }
        if ($pw) { $allSecrets["$($cfg.Tag)__DT_Password"] = $pw }

    }

    $cfg.DestinationPath = Select-FolderDialog "Папка, куда сохранять резервные копии"
    $cfg.Keep = [int](Read-Host "Сколько последних копий хранить (число)")

    # Яндекс.Диск (токен в секреты конкретной базы)
    $useCloud = Read-Choice "Отправлять копии в Яндекс.Диск?" @('Да','Нет')
    if ($useCloud -eq 1) {
        $cfg.CloudType = 'Yandex.Disk'
        $token = Read-Host "Введите OAuth токен Яндекс.Диска"
        if ($token) {
            $secretKey = "$($cfg.Tag)__YADiskToken"
            $allSecrets[$secretKey] = $token

            $currentCloudKeep = 0
            if ($cfg.PSObject.Properties.Name -contains 'CloudKeep') {
                try { $currentCloudKeep = [int]$cfg.CloudKeep } catch { $currentCloudKeep = 0 }
            } else {
                $cfg | Add-Member -NotePropertyName 'CloudKeep' -NotePropertyValue 0 -Force
            }
            $cloudKeepInput = Read-Host ("Сколько копий хранить в облаке (0 — без чистки), текущее: {0}" -f $currentCloudKeep)
            if (![string]::IsNullOrWhiteSpace($cloudKeepInput) -and [int]::TryParse($cloudKeepInput, [ref]([int]$null))) {
                $cfg.CloudKeep = [math]::Max(0, [int]$cloudKeepInput)
            } elseif ([string]::IsNullOrWhiteSpace($cloudKeepInput)) {
                $cfg.CloudKeep = $currentCloudKeep
            } else {
                Write-Host "Введено не число. CloudKeep оставлен равным {0}." -f $currentCloudKeep -ForegroundColor Yellow
                $cfg.CloudKeep = $currentCloudKeep
            }

            $cloudModulePath = Join-Path $PSScriptRoot '..\modules\Cloud.YandexDisk.psm1'
            $cloudModuleLoaded = $false
            if (Test-Path $cloudModulePath) {
                try {
                    Import-Module -Force -DisableNameChecking $cloudModulePath -ErrorAction Stop
                    $cloudModuleLoaded = $true
                }
                catch {
                    Write-Host ("Не удалось загрузить модуль Cloud.YandexDisk: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Файл modules\Cloud.YandexDisk.psm1 не найден." -ForegroundColor Yellow
            }

            if ($cloudModuleLoaded) {
                $testChoice = Read-Choice "Отправить тестовый файл на Яндекс.Диск для проверки?" @('Да','Нет')
                if ($testChoice -eq 1) {
                    $tmpFile = [IO.Path]::GetTempFileName()
                    try {
                        [IO.File]::WriteAllText($tmpFile, "Проверка связи с Яндекс.Диском {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
                        $remoteFolder = "/Backups1C/$($cfg.Tag)"
                        $remoteTest   = "$remoteFolder/__test_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date)
                        try {
                            Ensure-YandexDiskFolder -Token $token -RemotePath $remoteFolder
                            Upload-ToYandexDisk -Token $token -LocalPath $tmpFile -RemotePath $remoteTest -BarWidth 10
                            Write-Host "Тестовый файл успешно загружен (будет удалён)." -ForegroundColor Green
                            $headers = @{ Authorization = "OAuth $token" }
                            $encPath = [Uri]::EscapeDataString($remoteTest)
                            try {
                                Invoke-RestMethod -Uri "https://cloud-api.yandex.net/v1/disk/resources?path=$encPath&permanently=true" -Headers $headers -Method Delete -ErrorAction Stop | Out-Null
                            } catch { }
                        }
                        catch {
                            Write-Host ("Не удалось выгрузить тестовый файл: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                            $decision = Read-Choice "Как поступить?" @('Отключить облако для этой базы','Оставить включённым')
                            if ($decision -eq 1) {
                                $cfg.CloudType = ''
                                if ($allSecrets.ContainsKey($secretKey)) { $allSecrets.Remove($secretKey) }
                            }
                        }
                    }
                    finally {
                        Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            else {
                $decision = Read-Choice "Не удалось подготовить модуль облака. Отключить облако для этой базы?" @('Да','Нет')
                if ($decision -eq 1) {
                    $cfg.CloudType = ''
                    if ($allSecrets.ContainsKey($secretKey)) { $allSecrets.Remove($secretKey) }
                }
            }
        }
        else {
            Write-Host "Токен не указан — облако отключено." -ForegroundColor Yellow
            $cfg.CloudType = ''
        }
    } else {
        $cfg.CloudType = ''
    }

    $allConfigs += $cfg
    $more = Read-Choice "Добавить ещё одну базу?" @('Да','Нет')
    if ($more -ne 1) { break }
}

# ---------- telegram notifications ----------
Write-Host ''
$stateText = if ($telegramOptions.Enabled) { 'включена' } else { 'отключена' }
Write-Host ("Текущий статус Telegram-уведомлений: {0}" -f $stateText) -ForegroundColor Cyan

# Исправленный выбор 1/2
$enableChoice = Read-Choice "Включить отправку отчётов в Telegram?" @('Да', 'Нет')

$telegramTokenValue = $existingTelegramToken
$secondaryTokenValue = $existingTelegramSecondaryToken
$secondaryChatValue = $telegramOptions.SecondaryChatId

if ($enableChoice -eq 1) {
    $telegramOptions.Enabled = $true

    $tokenInput = Read-Host "Введите токен бота Telegram (оставьте пустым, чтобы сохранить текущее значение)"
    if (-not [string]::IsNullOrWhiteSpace($tokenInput)) { $telegramTokenValue = $tokenInput.Trim() }

    if (-not $telegramTokenValue) {
        Write-Host "Токен обязателен для отправки сообщений. Уведомления будут отключены." -ForegroundColor Yellow
        $telegramOptions.Enabled = $false
    }

    if ($telegramOptions.Enabled) {
        $chatPrompt = "Введите ID чата для отчётов (оставьте пустым, чтобы сохранить текущее значение)"
        if ([string]::IsNullOrWhiteSpace($telegramOptions.ChatId)) { $chatPrompt = "Введите ID чата для отчётов" }
        $chatInput = Read-Host $chatPrompt
        if (-not [string]::IsNullOrWhiteSpace($chatInput)) { $telegramOptions.ChatId = $chatInput.Trim() }

        if (-not $telegramOptions.ChatId) {
            Write-Host "ID чата обязателен. Уведомления будут отключены." -ForegroundColor Yellow
            $telegramOptions.Enabled = $false
        }
    }

    if ($telegramOptions.Enabled) {
        $modeChoice = Read-Choice "Отправлять отчёт только при ошибках?" @('Нет','Да')
        $telegramOptions.NotifyOnlyOnErrors = ($modeChoice -eq 2)

        # Второй бот
        $secondBotChoice = Read-Choice "Настроить второй бот для ответственного?" @('Нет', 'Да')
        if ($secondBotChoice -eq 2) {
            $secondaryTokenInput = Read-Host "Введите токен второго бота (оставьте пустым, чтобы сохранить текущее значение)"
            if (-not [string]::IsNullOrWhiteSpace($secondaryTokenInput)) {
                $secondaryTokenValue = $secondaryTokenInput.Trim()
            }

            if ($secondaryTokenValue) {
                $secondaryPrompt = "Введите ID чата второго бота (оставьте пустым, чтобы сохранить текущее значение)"
                if ([string]::IsNullOrWhiteSpace($secondaryChatValue)) { $secondaryPrompt = "Введите ID чата второго бота" }
                $secondaryInput = Read-Host $secondaryPrompt
                if (-not [string]::IsNullOrWhiteSpace($secondaryInput)) { $secondaryChatValue = $secondaryInput.Trim() }

                if (-not $secondaryChatValue) {
                    Write-Host "ID чата второго бота обязателен. Дополнительный бот отключён." -ForegroundColor Yellow
                    $secondaryTokenValue = ''
                }
            }
            else {
                $secondaryChatValue = ''
            }
        }
        else {
            $secondaryTokenValue = ''
            $secondaryChatValue = ''
        }
        
        # Тест Telegram
        $telegramModulePath = Join-Path $PSScriptRoot '..\modules\Notifications.Telegram.psm1'
        if (Test-Path $telegramModulePath) {
            try {
                Import-Module -Force -DisableNameChecking $telegramModulePath -ErrorAction Stop
                
                $testTgChoice = Read-Choice "Отправить тестовое сообщение в Telegram?" @('Да', 'Нет')
                if ($testTgChoice -eq 1) {
                    $testMessage = "Тест уведомлений от системы резервного копирования 1С`n" +
                                   "Дата и время: {0}`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') +
                                   "Статус: Настройка успешно завершена"
                    
                    try {
                        Send-TelegramMessage -Token $telegramTokenValue -ChatId $telegramOptions.ChatId -Text $testMessage
                        Write-Host "Тестовое сообщение успешно отправлено в основной чат." -ForegroundColor Green
                        
                        if ($secondaryTokenValue -and $secondaryChatValue) {
                            try {
                                Send-TelegramMessage -Token $secondaryTokenValue -ChatId $secondaryChatValue -Text $testMessage
                                Write-Host "Тестовое сообщение успешно отправлено во второй чат." -ForegroundColor Green
                            }
                            catch {
                                Write-Host ("Не удалось отправить сообщение во второй чат: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                            }
                        }
                    }
                    catch {
                        Write-Host ("Не удалось отправить тестовое сообщение: {0}" -f $_.Exception.Message) -ForegroundColor Red
                        $keepChoice = Read-Choice "Оставить настройки Telegram включёнными?" @('Да', 'Нет')
                        if ($keepChoice -eq 2) {
                            $telegramOptions.Enabled = $false
                            Write-Host "Telegram уведомления отключены." -ForegroundColor Yellow
                        }
                    }
                }
            }
            catch {
                Write-Host ("Не удалось загрузить модуль Telegram: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
        }
    }
    else {
        $telegramOptions.NotifyOnlyOnErrors = $false
        $telegramOptions.ChatId = ''
        $telegramTokenValue = ''
        $secondaryTokenValue = ''
        $secondaryChatValue = ''
    }
}
else {
    $telegramOptions.Enabled = $false
    $telegramOptions.ChatId = ''
    $telegramOptions.NotifyOnlyOnErrors = $false
    $secondaryTokenValue = ''
    $secondaryChatValue = ''
    $telegramTokenValue = ''
}

if ($telegramOptions.Enabled) {
    $allSecrets[$telegramTokenKey] = $telegramTokenValue
    if ($secondaryTokenValue) {
        $allSecrets[$secondaryTokenKey] = $secondaryTokenValue
    }
    elseif ($allSecrets.ContainsKey($secondaryTokenKey)) {
        $allSecrets.Remove($secondaryTokenKey)
        $secondaryChatValue = ''
    }
}
else {
    if ($allSecrets.ContainsKey($telegramTokenKey)) { $allSecrets.Remove($telegramTokenKey) }
    if ($allSecrets.ContainsKey($secondaryTokenKey)) { $allSecrets.Remove($secondaryTokenKey) }
}

$telegramOptions.SecondaryChatId = if ($telegramOptions.Enabled) { $secondaryChatValue } else { '' }

# ---------- save ----------
foreach ($cfg in $allConfigs) {
    $cfgPath = Join-Path $basesDir ("{0}.json" -f $cfg.Tag)
    $cfg | ConvertTo-Json -Depth 8 | Set-Content -Path $cfgPath -Encoding UTF8
    Write-Host ("Готово. База [{0}] добавлена." -f $cfg.Tag) -ForegroundColor Green
}

Encrypt-Secrets -Secrets $allSecrets -KeyPath $keyPath -OutFile $secretsFile
Write-Host "Секреты сохранены и зашифрованы." -ForegroundColor Green

$act = Read-Choice "Что делать после завершения бэкапа?" @('Выключить ПК','Перезагрузить ПК','Ничего не делать')
if ($telegramOptions.ChatId) { $telegramOptions.ChatId = $telegramOptions.ChatId.Trim() }
$settingsToSave = if ($settingsData -is [hashtable]) { [hashtable]$settingsData.Clone() } else { @{} }
$settingsToSave['AfterBackup'] = $act
$settingsToSave['Telegram'] = $telegramOptions
$settingsToSave | ConvertTo-Json -Depth 6 | Set-Content -Path $settingsFile -Encoding UTF8

Write-Host "[INFO] Настройка баз завершена." -ForegroundColor Green