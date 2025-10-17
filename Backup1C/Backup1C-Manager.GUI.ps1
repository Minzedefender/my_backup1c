#requires -Version 5.1
# Объединенный менеджер системы резервного копирования 1С
# Включает функционал мастера настройки и редактора конфигураций

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Пути и переменные
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = $ScriptDir
$ConfigRoot = Join-Path $ProjectRoot 'config'
$BasesDir = Join-Path $ConfigRoot 'bases'
$SettingsFile = Join-Path $ConfigRoot 'settings.json'
$KeyPath = Join-Path $ConfigRoot 'key.bin'
$SecretsFile = Join-Path $ConfigRoot 'secrets.json.enc'

# Создание директорий
if (-not (Test-Path $ConfigRoot)) { New-Item -ItemType Directory -Path $ConfigRoot -Force | Out-Null }
if (-not (Test-Path $BasesDir)) { New-Item -ItemType Directory -Path $BasesDir -Force | Out-Null }

# Импорт модулей
try {
    Import-Module -Force -DisableNameChecking (Join-Path $ProjectRoot 'modules\Common.Crypto.psm1') -ErrorAction Stop
    Import-Module -Force -DisableNameChecking (Join-Path $ProjectRoot 'modules\System.Scheduler.psm1') -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show("Не удалось загрузить модули: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# Глобальные переменные
$Script:AllBases = @()
$Script:AllSecrets = @{}
$Script:CurrentBase = $null
$Script:HasChanges = $false


# === ГЛАВНАЯ ФОРМА ===
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Менеджер системы резервного копирования 1С"
$MainForm.Size = New-Object System.Drawing.Size(1200, 900)
$MainForm.StartPosition = "CenterScreen"
$MainForm.MinimumSize = New-Object System.Drawing.Size(1200, 900)
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# === TABCONTROL ===
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Point(12, 12)
$TabControl.Size = New-Object System.Drawing.Size(1160, 840)
$TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# === ВКЛАДКА МАСТЕР НАСТРОЙКИ ===
$TabWizard = New-Object System.Windows.Forms.TabPage
$TabWizard.Text = "Мастер настройки"
$TabWizard.UseVisualStyleBackColor = $true

# === ВКЛАДКА РЕДАКТОР ===
$TabEditor = New-Object System.Windows.Forms.TabPage
$TabEditor.Text = "Редактор конфигураций"
$TabEditor.UseVisualStyleBackColor = $true

# === ВКЛАДКА РАСПИСАНИЕ ===
$TabScheduler = New-Object System.Windows.Forms.TabPage
$TabScheduler.Text = "Расписание"
$TabScheduler.UseVisualStyleBackColor = $true

# === ВКЛАДКА МОНИТОРИНГ ===
$TabMonitor = New-Object System.Windows.Forms.TabPage
$TabMonitor.Text = "Мониторинг"
$TabMonitor.UseVisualStyleBackColor = $true

$TabControl.TabPages.Add($TabWizard)
$TabControl.TabPages.Add($TabEditor)
$TabControl.TabPages.Add($TabScheduler)
$TabControl.TabPages.Add($TabMonitor)
$MainForm.Controls.Add($TabControl)

# === ФУНКЦИИ ОБЩЕГО НАЗНАЧЕНИЯ ===
function ConvertTo-Hashtable($obj) {
    if ($obj -is [hashtable]) { return $obj }
    $ht = @{}
    if ($obj -is [System.Collections.IDictionary]) {
        foreach($k in $obj.Keys){ $ht[$k] = $obj[$k] }
    } elseif ($obj -and $obj.PSObject) {
        foreach($p in $obj.PSObject.Properties){ $ht[$p.Name] = $p.Value }
    }
    return $ht
}

function Load-ExistingSettings {
    # Загрузка секретов
    if ((Test-Path $SecretsFile) -and (Test-Path $KeyPath)) {
        try {
            $raw = Decrypt-Secrets -InFile $SecretsFile -KeyPath $KeyPath
            $Script:AllSecrets = ConvertTo-Hashtable $raw
            if (-not ($Script:AllSecrets -is [hashtable])) { $Script:AllSecrets = @{} }
        } catch { $Script:AllSecrets = @{} }
    }

    # Загрузка настроек
    $settingsData = @{}
    if (Test-Path $SettingsFile) {
        try {
            $settingsData = ConvertTo-Hashtable (Get-Content $SettingsFile -Raw | ConvertFrom-Json)
            if (-not ($settingsData -is [hashtable])) { $settingsData = @{} }
        } catch { $settingsData = @{} }
    }

    return $settingsData
}

function Get-WebServices {
    $masks = @('Apache2.4','Apache*','httpd*','W3SVC','WAS')
    $all = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $n = $_.Name
        $masks | Where-Object { $n -like $_ } | Select-Object -First 1
    }
    return $all | Sort-Object @{Expression='Status';Descending=$true}, @{Expression='DisplayName';Descending=$false}
}

function Get-1CEStartCandidates {
    $list = @()
    $pf = $env:ProgramFiles
    $pf86 = ${env:ProgramFiles(x86)}
    if ($pf) { $list += (Join-Path $pf '1cv8\common\1cestart.exe') }
    if ($pf86) { $list += (Join-Path $pf86 '1cv8\common\1cestart.exe') }
    return $list | Where-Object { $_ -and (Test-Path $_) }
}

# ==========================================
# МАСТЕР НАСТРОЙКИ
# ==========================================

# Переменные мастера
$Script:CurrentStep = 1
$Script:MaxSteps = 3

# === ПАНЕЛИ МАСТЕРА ===
$WizardProgressPanel = New-Object System.Windows.Forms.Panel
$WizardProgressPanel.Location = New-Object System.Drawing.Point(12, 12)
$WizardProgressPanel.Size = New-Object System.Drawing.Size(925, 60)
$WizardProgressPanel.BackColor = [System.Drawing.Color]::WhiteSmoke

$WizardProgressLabel = New-Object System.Windows.Forms.Label
$WizardProgressLabel.Location = New-Object System.Drawing.Point(10, 10)
$WizardProgressLabel.Size = New-Object System.Drawing.Size(900, 20)
$WizardProgressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$WizardProgressLabel.Text = "Шаг 1 из 3: Настройка баз данных"

$WizardProgressBar = New-Object System.Windows.Forms.ProgressBar
$WizardProgressBar.Location = New-Object System.Drawing.Point(10, 35)
$WizardProgressBar.Size = New-Object System.Drawing.Size(900, 20)
$WizardProgressBar.Maximum = $Script:MaxSteps
$WizardProgressBar.Value = $Script:CurrentStep

$WizardProgressPanel.Controls.AddRange(@($WizardProgressLabel, $WizardProgressBar))

$WizardContentPanel = New-Object System.Windows.Forms.Panel
$WizardContentPanel.Location = New-Object System.Drawing.Point(12, 85)
$WizardContentPanel.Size = New-Object System.Drawing.Size(925, 480)
$WizardContentPanel.BackColor = [System.Drawing.Color]::White
$WizardContentPanel.BorderStyle = "FixedSingle"

$WizardButtonPanel = New-Object System.Windows.Forms.Panel
$WizardButtonPanel.Location = New-Object System.Drawing.Point(12, 575)
$WizardButtonPanel.Size = New-Object System.Drawing.Size(925, 50)

$WizardButtonBack = New-Object System.Windows.Forms.Button
$WizardButtonBack.Text = "< Назад"
$WizardButtonBack.Location = New-Object System.Drawing.Point(10, 10)
$WizardButtonBack.Size = New-Object System.Drawing.Size(100, 35)
$WizardButtonBack.Enabled = $false

$WizardButtonNext = New-Object System.Windows.Forms.Button
$WizardButtonNext.Text = "Далее >"
$WizardButtonNext.Location = New-Object System.Drawing.Point(710, 10)
$WizardButtonNext.Size = New-Object System.Drawing.Size(100, 35)
$WizardButtonNext.UseVisualStyleBackColor = $true

$WizardButtonFinish = New-Object System.Windows.Forms.Button
$WizardButtonFinish.Text = "Завершить"
$WizardButtonFinish.Location = New-Object System.Drawing.Point(820, 10)
$WizardButtonFinish.Size = New-Object System.Drawing.Size(100, 35)
$WizardButtonFinish.BackColor = [System.Drawing.Color]::LightGreen
$WizardButtonFinish.UseVisualStyleBackColor = $false
$WizardButtonFinish.Visible = $false

$WizardButtonPanel.Controls.AddRange(@($WizardButtonBack, $WizardButtonNext, $WizardButtonFinish))
$TabWizard.Controls.AddRange(@($WizardProgressPanel, $WizardContentPanel, $WizardButtonPanel))

# === ФУНКЦИИ МАСТЕРА ===
function Show-WizardStep1 {
    $WizardContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Настройка баз данных для резервного копирования"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(880, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Список баз
    $BasesGroupBox = New-Object System.Windows.Forms.GroupBox
    $BasesGroupBox.Text = "Список баз для резервного копирования"
    $BasesGroupBox.Location = New-Object System.Drawing.Point(20, 60)
    $BasesGroupBox.Size = New-Object System.Drawing.Size(880, 300)

    $Script:WizardBasesListView = New-Object System.Windows.Forms.ListView
    $Script:WizardBasesListView.View = "Details"
    $Script:WizardBasesListView.FullRowSelect = $true
    $Script:WizardBasesListView.GridLines = $true
    $Script:WizardBasesListView.Location = New-Object System.Drawing.Point(10, 25)
    $Script:WizardBasesListView.Size = New-Object System.Drawing.Size(860, 230)

    [void]$Script:WizardBasesListView.Columns.Add("Имя базы", 150)
    [void]$Script:WizardBasesListView.Columns.Add("Тип", 60)
    [void]$Script:WizardBasesListView.Columns.Add("Путь к базе", 280)
    [void]$Script:WizardBasesListView.Columns.Add("Папка бэкапов", 230)
    [void]$Script:WizardBasesListView.Columns.Add("Копий", 60)
    [void]$Script:WizardBasesListView.Columns.Add("Облако", 80)

    $AddBaseButton = New-Object System.Windows.Forms.Button
    $AddBaseButton.Text = "Добавить базу"
    $AddBaseButton.Location = New-Object System.Drawing.Point(10, 265)
    $AddBaseButton.Size = New-Object System.Drawing.Size(120, 30)
    $AddBaseButton.BackColor = [System.Drawing.Color]::LightGreen

    $DeleteBaseButton = New-Object System.Windows.Forms.Button
    $DeleteBaseButton.Text = "Удалить"
    $DeleteBaseButton.Location = New-Object System.Drawing.Point(140, 265)
    $DeleteBaseButton.Size = New-Object System.Drawing.Size(120, 30)
    $DeleteBaseButton.BackColor = [System.Drawing.Color]::LightCoral
    $DeleteBaseButton.Enabled = $false

    $BasesGroupBox.Controls.AddRange(@($Script:WizardBasesListView, $AddBaseButton, $DeleteBaseButton))

    # Информационная панель
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Text = "Добавьте базы данных 1С, для которых нужно настроить резервное копирование.`nВы можете настроить копирование файлов .1CD или выгрузку в формат .dt через конфигуратор.`n`nДля редактирования базы дважды кликните на неё в списке."
    $InfoLabel.Location = New-Object System.Drawing.Point(20, 380)
    $InfoLabel.Size = New-Object System.Drawing.Size(880, 70)
    $InfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $InfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen

    $WizardContentPanel.Controls.AddRange(@($TitleLabel, $BasesGroupBox, $InfoLabel))

    # Обработчики событий
    $Script:WizardBasesListView.Add_SelectedIndexChanged({
        $hasSelection = $Script:WizardBasesListView.SelectedItems.Count -gt 0
        $DeleteBaseButton.Enabled = $hasSelection
    })

    $Script:WizardBasesListView.Add_DoubleClick({
        if ($Script:WizardBasesListView.SelectedItems.Count -gt 0) {
            Show-WizardBaseEditDialog -EditMode $true
        }
    })

    $AddBaseButton.Add_Click({ Show-WizardBaseEditDialog })
    $DeleteBaseButton.Add_Click({
        if ($Script:WizardBasesListView.SelectedItems.Count -gt 0) {
            $selectedIndex = $Script:WizardBasesListView.SelectedItems[0].Index
            $baseName = $Script:AllBases[$selectedIndex].Tag
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Удалить базу '$baseName' из настроек?",
                "Подтверждение",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Удаляем секреты базы
                $keysToRemove = @()
                foreach ($key in $Script:AllSecrets.Keys) {
                    if ($key.StartsWith("$baseName" + "__")) {
                        $keysToRemove += $key
                    }
                }
                foreach ($key in $keysToRemove) {
                    $Script:AllSecrets.Remove($key)
                }

                # Удаляем базу из списка
                $Script:AllBases = $Script:AllBases | Where-Object { $_.Tag -ne $baseName }
                Update-WizardBasesListView

                # Сохраняем обновленные секреты
                try {
                    Encrypt-Secrets -Secrets $Script:AllSecrets -KeyPath $KeyPath -OutFile $SecretsFile
                } catch {
                    # Игнорируем ошибки сохранения секретов при удалении
                }
            }
        }
    })

    Update-WizardBasesListView
}

function Update-WizardBasesListView {
    $Script:WizardBasesListView.Items.Clear()
    foreach ($base in $Script:AllBases) {
        $item = New-Object System.Windows.Forms.ListViewItem($base.Tag)
        [void]$item.SubItems.Add($base.BackupType)
        [void]$item.SubItems.Add($base.SourcePath)
        [void]$item.SubItems.Add($base.DestinationPath)
        [void]$item.SubItems.Add($base.Keep.ToString())
        $cloudText = if ($base.CloudType -eq 'Yandex.Disk') { 'Да' } else { 'Нет' }
        [void]$item.SubItems.Add($cloudText)
        [void]$Script:WizardBasesListView.Items.Add($item)
    }
}

function Show-WizardBaseEditDialog {
    param([bool]$EditMode = $false)

    $editBase = $null
    $editIndex = -1

    if ($EditMode -and $Script:WizardBasesListView.SelectedItems.Count -gt 0) {
        $editIndex = $Script:WizardBasesListView.SelectedItems[0].Index
        $editBase = $Script:AllBases[$editIndex]
    }

    $BaseDialog = New-Object System.Windows.Forms.Form
    $BaseDialog.Text = if ($EditMode) { "Редактирование базы" } else { "Добавление новой базы" }
    $BaseDialog.Size = New-Object System.Drawing.Size(600, 900)
    $BaseDialog.StartPosition = "CenterParent"
    $BaseDialog.FormBorderStyle = "FixedDialog"
    $BaseDialog.MaximizeBox = $false
    $BaseDialog.MinimizeBox = $false

    # Имя базы
    $LabelName = New-Object System.Windows.Forms.Label
    $LabelName.Text = "Уникальное имя базы:"
    $LabelName.Location = New-Object System.Drawing.Point(15, 20)
    $LabelName.Size = New-Object System.Drawing.Size(200, 20)

    $TextName = New-Object System.Windows.Forms.TextBox
    $TextName.Location = New-Object System.Drawing.Point(15, 45)
    $TextName.Size = New-Object System.Drawing.Size(550, 25)
    if ($editBase) { $TextName.Text = $editBase.Tag }


    # Тип бэкапа
    $LabelType = New-Object System.Windows.Forms.Label
    $LabelType.Text = "Тип резервной копии:"
    $LabelType.Location = New-Object System.Drawing.Point(15, 80)
    $LabelType.Size = New-Object System.Drawing.Size(200, 20)

    $ComboType = New-Object System.Windows.Forms.ComboBox
    $ComboType.Items.AddRange(@("1CD", "DT"))
    $ComboType.DropDownStyle = "DropDownList"
    $ComboType.Location = New-Object System.Drawing.Point(15, 105)
    $ComboType.Size = New-Object System.Drawing.Size(550, 25)
    if ($editBase) { $ComboType.Text = $editBase.BackupType } else { $ComboType.SelectedIndex = 0 }


    # Путь к базе
    $LabelSource = New-Object System.Windows.Forms.Label
    $LabelSource.Text = "Путь к базе данных:"
    $LabelSource.Location = New-Object System.Drawing.Point(15, 140)
    $LabelSource.Size = New-Object System.Drawing.Size(200, 20)

    $TextSource = New-Object System.Windows.Forms.TextBox
    $TextSource.Location = New-Object System.Drawing.Point(15, 165)
    $TextSource.Size = New-Object System.Drawing.Size(450, 25)

    if ($editBase) { $TextSource.Text = $editBase.SourcePath }

    $ButtonBrowseSource = New-Object System.Windows.Forms.Button
    $ButtonBrowseSource.Text = "Обзор..."
    $ButtonBrowseSource.Location = New-Object System.Drawing.Point(475, 164)
    $ButtonBrowseSource.Size = New-Object System.Drawing.Size(90, 27)

    # Папка для бэкапов
    $LabelDest = New-Object System.Windows.Forms.Label
    $LabelDest.Text = "Папка для резервных копий:"
    $LabelDest.Location = New-Object System.Drawing.Point(15, 200)
    $LabelDest.Size = New-Object System.Drawing.Size(200, 20)

    $TextDest = New-Object System.Windows.Forms.TextBox
    $TextDest.Location = New-Object System.Drawing.Point(15, 225)
    $TextDest.Size = New-Object System.Drawing.Size(450, 25)
    if ($editBase) { $TextDest.Text = $editBase.DestinationPath }


    $ButtonBrowseDest = New-Object System.Windows.Forms.Button
    $ButtonBrowseDest.Text = "Обзор..."
    $ButtonBrowseDest.Location = New-Object System.Drawing.Point(475, 224)
    $ButtonBrowseDest.Size = New-Object System.Drawing.Size(90, 27)

    # Количество копий
    $LabelKeep = New-Object System.Windows.Forms.Label
    $LabelKeep.Text = "Хранить копий локально:"
    $LabelKeep.Location = New-Object System.Drawing.Point(15, 260)
    $LabelKeep.Size = New-Object System.Drawing.Size(200, 20)

    $NumericKeep = New-Object System.Windows.Forms.NumericUpDown
    $NumericKeep.Location = New-Object System.Drawing.Point(15, 285)
    $NumericKeep.Size = New-Object System.Drawing.Size(100, 25)
    $NumericKeep.Minimum = 1
    $NumericKeep.Maximum = 999
    if ($editBase) { $NumericKeep.Value = $editBase.Keep } else { $NumericKeep.Value = 3 }


    # Настройки для DT
    $GroupBoxDT = New-Object System.Windows.Forms.GroupBox
    $GroupBoxDT.Text = "Настройки для выгрузки DT"
    $GroupBoxDT.Location = New-Object System.Drawing.Point(15, 320)
    $GroupBoxDT.Size = New-Object System.Drawing.Size(550, 230)
    $GroupBoxDT.Enabled = ($ComboType.Text -eq "DT")

    # 1cestart.exe
    $LabelExe = New-Object System.Windows.Forms.Label
    $LabelExe.Text = "Путь к 1cestart.exe (обязательно для DT):"
    $LabelExe.Location = New-Object System.Drawing.Point(10, 25)
    $LabelExe.Size = New-Object System.Drawing.Size(250, 20)
    $LabelExe.ForeColor = [System.Drawing.Color]::DarkRed

    $TextExe = New-Object System.Windows.Forms.TextBox
    $TextExe.Location = New-Object System.Drawing.Point(10, 50)
    $TextExe.Size = New-Object System.Drawing.Size(420, 25)
    if ($editBase -and $editBase.PSObject.Properties.Name -contains 'ExePath') { $TextExe.Text = $editBase.ExePath }


    $ButtonBrowseExe = New-Object System.Windows.Forms.Button
    $ButtonBrowseExe.Text = "Обзор..."
    $ButtonBrowseExe.Location = New-Object System.Drawing.Point(440, 49)
    $ButtonBrowseExe.Size = New-Object System.Drawing.Size(90, 27)

    # Логин и пароль для DT
    $LabelDTLogin = New-Object System.Windows.Forms.Label
    $LabelDTLogin.Text = "Логин 1С (можно оставить пустым):"
    $LabelDTLogin.Location = New-Object System.Drawing.Point(10, 85)
    $LabelDTLogin.Size = New-Object System.Drawing.Size(250, 20)

    $TextDTLogin = New-Object System.Windows.Forms.TextBox
    $TextDTLogin.Location = New-Object System.Drawing.Point(10, 110)
    $TextDTLogin.Size = New-Object System.Drawing.Size(200, 25)
    if ($editBase) {
        $loginKey = "$($editBase.Tag)__DT_Login"
        if ($Script:AllSecrets.ContainsKey($loginKey)) {
            $TextDTLogin.Text = [string]$Script:AllSecrets[$loginKey]
        }
    }


    $LabelDTPassword = New-Object System.Windows.Forms.Label
    $LabelDTPassword.Text = "Пароль 1С (можно оставить пустым):"
    $LabelDTPassword.Location = New-Object System.Drawing.Point(220, 85)
    $LabelDTPassword.Size = New-Object System.Drawing.Size(250, 20)

    $TextDTPassword = New-Object System.Windows.Forms.TextBox
    $TextDTPassword.Location = New-Object System.Drawing.Point(220, 110)
    $TextDTPassword.Size = New-Object System.Drawing.Size(200, 25)
    $TextDTPassword.UseSystemPasswordChar = $true
    if ($editBase) {
        $passwordKey = "$($editBase.Tag)__DT_Password"
        if ($Script:AllSecrets.ContainsKey($passwordKey)) {
            $TextDTPassword.Text = [string]$Script:AllSecrets[$passwordKey]
        }
    }

    # Службы для остановки
    $LabelServices = New-Object System.Windows.Forms.Label
    $LabelServices.Text = "Службы для остановки:"
    $LabelServices.Location = New-Object System.Drawing.Point(10, 145)
    $LabelServices.Size = New-Object System.Drawing.Size(200, 20)

    $RadioAutoServices = New-Object System.Windows.Forms.RadioButton
    $RadioAutoServices.Text = "Автоматически (Apache2.4 или все веб-службы)"
    $RadioAutoServices.Location = New-Object System.Drawing.Point(10, 170)
    $RadioAutoServices.Size = New-Object System.Drawing.Size(350, 20)
    $RadioAutoServices.Checked = $true

    $RadioManualServices = New-Object System.Windows.Forms.RadioButton
    $RadioManualServices.Text = "Выбрать вручную"
    $RadioManualServices.Location = New-Object System.Drawing.Point(370, 170)
    $RadioManualServices.Size = New-Object System.Drawing.Size(150, 20)

    $RadioNoServices = New-Object System.Windows.Forms.RadioButton
    $RadioNoServices.Text = "Не останавливать"
    $RadioNoServices.Location = New-Object System.Drawing.Point(10, 195)
    $RadioNoServices.Size = New-Object System.Drawing.Size(150, 20)

    # Определяем текущий выбор на основе существующих настроек
    if ($editBase -and $editBase.PSObject.Properties.Name -contains 'StopServices') {
        if ($editBase.StopServices.Count -eq 0) {
            $RadioNoServices.Checked = $true
            $RadioAutoServices.Checked = $false
        } elseif ($editBase.StopServices.Count -eq 1 -and $editBase.StopServices[0] -eq 'Apache2.4') {
            $RadioAutoServices.Checked = $true
        } else {
            $RadioManualServices.Checked = $true
            $RadioAutoServices.Checked = $false
        }
    }

    $GroupBoxDT.Controls.AddRange(@($LabelExe, $TextExe, $ButtonBrowseExe, $LabelDTLogin, $TextDTLogin, $LabelDTPassword, $TextDTPassword, $LabelServices, $RadioAutoServices, $RadioManualServices, $RadioNoServices))

    # Облачное хранилище
    $GroupBoxCloud = New-Object System.Windows.Forms.GroupBox
    $GroupBoxCloud.Text = "Яндекс.Диск"
    $GroupBoxCloud.Location = New-Object System.Drawing.Point(15, 560)
    $GroupBoxCloud.Size = New-Object System.Drawing.Size(550, 120)

    $CheckBoxUseCloud = New-Object System.Windows.Forms.CheckBox
    $CheckBoxUseCloud.Text = "Отправлять копии в Яндекс.Диск"
    $CheckBoxUseCloud.Location = New-Object System.Drawing.Point(10, 25)
    $CheckBoxUseCloud.Size = New-Object System.Drawing.Size(250, 20)
    if ($editBase) { $CheckBoxUseCloud.Checked = ($editBase.CloudType -eq 'Yandex.Disk') }


    $LabelCloudToken = New-Object System.Windows.Forms.Label
    $LabelCloudToken.Text = "OAuth токен Яндекс.Диска:"
    $LabelCloudToken.Location = New-Object System.Drawing.Point(10, 50)
    $LabelCloudToken.Size = New-Object System.Drawing.Size(200, 20)
    $LabelCloudToken.Enabled = $CheckBoxUseCloud.Checked


    $TextCloudToken = New-Object System.Windows.Forms.TextBox
    $TextCloudToken.Location = New-Object System.Drawing.Point(10, 75)
    $TextCloudToken.Size = New-Object System.Drawing.Size(300, 25)
    $TextCloudToken.UseSystemPasswordChar = $true
    $TextCloudToken.Enabled = $CheckBoxUseCloud.Checked
    if ($editBase) {
        $tokenKey = "$($editBase.Tag)__YADiskToken"
        if ($Script:AllSecrets.ContainsKey($tokenKey)) {
            $TextCloudToken.Text = [string]$Script:AllSecrets[$tokenKey]
        }
    }

    $LabelCloudKeep = New-Object System.Windows.Forms.Label
    $LabelCloudKeep.Text = "Копий в облаке:"
    $LabelCloudKeep.Location = New-Object System.Drawing.Point(320, 50)
    $LabelCloudKeep.Size = New-Object System.Drawing.Size(100, 20)
    $LabelCloudKeep.Enabled = $CheckBoxUseCloud.Checked

    $NumericCloudKeep = New-Object System.Windows.Forms.NumericUpDown
    $NumericCloudKeep.Location = New-Object System.Drawing.Point(320, 75)
    $NumericCloudKeep.Size = New-Object System.Drawing.Size(80, 25)
    $NumericCloudKeep.Minimum = 0
    $NumericCloudKeep.Maximum = 999
    $NumericCloudKeep.Value = if ($editBase -and $editBase.PSObject.Properties.Name -contains 'CloudKeep') { $editBase.CloudKeep } else { 3 }
    $NumericCloudKeep.Enabled = $CheckBoxUseCloud.Checked

    $ButtonTestCloud = New-Object System.Windows.Forms.Button
    $ButtonTestCloud.Text = "Тест подключения"
    $ButtonTestCloud.Location = New-Object System.Drawing.Point(420, 75)
    $ButtonTestCloud.Size = New-Object System.Drawing.Size(120, 25)
    $ButtonTestCloud.BackColor = [System.Drawing.Color]::LightBlue
    $ButtonTestCloud.Enabled = $CheckBoxUseCloud.Checked

    $GroupBoxCloud.Controls.AddRange(@($CheckBoxUseCloud, $LabelCloudToken, $TextCloudToken, $LabelCloudKeep, $NumericCloudKeep, $ButtonTestCloud))

    # === ГРУППА TELEGRAM НАСТРОЙКИ ===
    $GroupBoxTelegram = New-Object System.Windows.Forms.GroupBox
    $GroupBoxTelegram.Text = "Индивидуальные настройки Telegram"
    $GroupBoxTelegram.Location = New-Object System.Drawing.Point(15, 690)
    $GroupBoxTelegram.Size = New-Object System.Drawing.Size(550, 120)

    $CheckBoxTelegramOverride = New-Object System.Windows.Forms.CheckBox
    $CheckBoxTelegramOverride.Text = "Переопределить настройки Telegram для этой базы"
    $CheckBoxTelegramOverride.Location = New-Object System.Drawing.Point(10, 25)
    $CheckBoxTelegramOverride.Size = New-Object System.Drawing.Size(350, 20)
    if ($editBase -and $editBase.PSObject.Properties.Name -contains 'TelegramOverride') {
        $CheckBoxTelegramOverride.Checked = [bool]$editBase.TelegramOverride
    }

    $LabelTelegramChatId = New-Object System.Windows.Forms.Label
    $LabelTelegramChatId.Text = "Chat ID для уведомлений:"
    $LabelTelegramChatId.Location = New-Object System.Drawing.Point(10, 50)
    $LabelTelegramChatId.Size = New-Object System.Drawing.Size(150, 20)
    $LabelTelegramChatId.Enabled = $CheckBoxTelegramOverride.Checked

    $TextTelegramChatId = New-Object System.Windows.Forms.TextBox
    $TextTelegramChatId.Location = New-Object System.Drawing.Point(10, 75)
    $TextTelegramChatId.Size = New-Object System.Drawing.Size(200, 25)
    $TextTelegramChatId.Enabled = $CheckBoxTelegramOverride.Checked
    if ($editBase -and $editBase.PSObject.Properties.Name -contains 'TelegramChatId') {
        $TextTelegramChatId.Text = [string]$editBase.TelegramChatId
    }

    $CheckBoxTelegramOnlyErrors = New-Object System.Windows.Forms.CheckBox
    $CheckBoxTelegramOnlyErrors.Text = "Только при ошибках"
    $CheckBoxTelegramOnlyErrors.Location = New-Object System.Drawing.Point(220, 75)
    $CheckBoxTelegramOnlyErrors.Size = New-Object System.Drawing.Size(150, 25)
    $CheckBoxTelegramOnlyErrors.Enabled = $CheckBoxTelegramOverride.Checked
    if ($editBase -and $editBase.PSObject.Properties.Name -contains 'TelegramOnlyErrors') {
        $CheckBoxTelegramOnlyErrors.Checked = [bool]$editBase.TelegramOnlyErrors
    }

    $ButtonTestTelegram = New-Object System.Windows.Forms.Button
    $ButtonTestTelegram.Text = "Тест"
    $ButtonTestTelegram.Location = New-Object System.Drawing.Point(380, 73)
    $ButtonTestTelegram.Size = New-Object System.Drawing.Size(60, 27)
    $ButtonTestTelegram.Enabled = $CheckBoxTelegramOverride.Checked

    $GroupBoxTelegram.Controls.AddRange(@($CheckBoxTelegramOverride, $LabelTelegramChatId, $TextTelegramChatId, $CheckBoxTelegramOnlyErrors, $ButtonTestTelegram))

    # Кнопки
    $ButtonOK = New-Object System.Windows.Forms.Button
    $ButtonOK.Text = "OK"
    $ButtonOK.Location = New-Object System.Drawing.Point(400, 820)
    $ButtonOK.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $ButtonCancel = New-Object System.Windows.Forms.Button
    $ButtonCancel.Text = "Отмена"
    $ButtonCancel.Location = New-Object System.Drawing.Point(490, 820)
    $ButtonCancel.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $BaseDialog.Controls.AddRange(@($LabelName, $TextName, $LabelType, $ComboType, $LabelSource, $TextSource, $ButtonBrowseSource, $LabelDest, $TextDest, $ButtonBrowseDest, $LabelKeep, $NumericKeep, $GroupBoxDT, $GroupBoxCloud, $GroupBoxTelegram, $ButtonOK, $ButtonCancel))

    # Обработчики событий
    $ComboType.Add_SelectedIndexChanged({
        try {
            $GroupBoxDT.Enabled = ($ComboType.Text -eq "DT")
        }
        catch {
            # Игнорируем ошибки при инициализации
        }
    })

    # Telegram override checkbox handler
    $CheckBoxTelegramOverride.Add_CheckedChanged({
        $TextTelegramChatId.Enabled = $CheckBoxTelegramOverride.Checked
        $CheckBoxTelegramOnlyErrors.Enabled = $CheckBoxTelegramOverride.Checked
        $ButtonTestTelegram.Enabled = $CheckBoxTelegramOverride.Checked
    })

    # Test Telegram button handler
    $ButtonTestTelegram.Add_Click({
        if ($TextTelegramChatId.Text -and $TextTelegramChatId.Text.Trim() -ne "") {
            try {
                # Load global config for bot token
                $globalConfigPath = Join-Path $PSScriptRoot "config\global.json"
                if (Test-Path $globalConfigPath) {
                    $globalConfig = Get-Content $globalConfigPath | ConvertFrom-Json
                    if ($globalConfig.TelegramBotToken) {
                        $telegramConfig = @{
                            ChatId = $TextTelegramChatId.Text.Trim()
                            BotToken = $globalConfig.TelegramBotToken
                        }
                        Send-TelegramMessage -Config $telegramConfig -Message "Test message from Backup1C Manager"
                        [System.Windows.Forms.MessageBox]::Show("Test message sent successfully!", "Telegram Test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Bot token not found in global configuration!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Global configuration file not found!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error sending test message: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter ChatId first!", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $ButtonBrowseExe.Add_Click({
        $candidates = Get-1CEStartCandidates
        $initialDir = if ($candidates -and $candidates[0]) { Split-Path $candidates[0] -Parent } else { "$env:ProgramFiles\1cv8\common" }

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "1cestart.exe|1cestart.exe|Все exe файлы (*.exe)|*.exe"
        $dialog.Title = "Выберите 1cestart.exe"
        $dialog.InitialDirectory = $initialDir
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $TextExe.Text = $dialog.FileName
        }
    })

    # Обработчик для включения/отключения полей Яндекс.Диска
    $CheckBoxUseCloud.Add_CheckedChanged({
        try {
            $enabled = $CheckBoxUseCloud.Checked
            $LabelCloudToken.Enabled = $enabled
            $TextCloudToken.Enabled = $enabled
            $LabelCloudKeep.Enabled = $enabled
            $NumericCloudKeep.Enabled = $enabled
            $ButtonTestCloud.Enabled = $enabled
        }
        catch {
            # Игнорируем ошибки при инициализации
        }
    })

    # Обработчик для теста Яндекс.Диска
    $ButtonTestCloud.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($TextCloudToken.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Введите токен Яндекс.Диска!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $baseName = if ([string]::IsNullOrWhiteSpace($TextName.Text)) { 'Test' } else { $TextName.Text }
            Test-YandexDiskConnection -Token $TextCloudToken.Text -BaseName $baseName
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Обработчик для выбора служб вручную
    $RadioManualServices.Add_CheckedChanged({
        if ($RadioManualServices.Checked) {
            $selectedServices = Show-ServicesSelectionDialog
            if ($selectedServices -and $selectedServices.Count -gt 0) {
                # Сохраним выбранные службы для последующего использования
                $Script:SelectedServices = $selectedServices
            } else {
                # Если ничего не выбрано, переключаемся на "Не останавливать"
                $RadioNoServices.Checked = $true
                $RadioManualServices.Checked = $false
            }
        }
    })

    # Обработчики
    $ButtonBrowseSource.Add_Click({
        if ($ComboType.Text -eq "1CD") {
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = "База 1С (*.1CD)|*.1CD|Все файлы (*.*)|*.*"
            $dialog.Title = "Выберите файл базы 1С"
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $TextSource.Text = $dialog.FileName
            }
        } else {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Выберите папку с базой 1С"
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $TextSource.Text = $dialog.SelectedPath
            }
        }
    })

    $ButtonBrowseDest.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Выберите папку для резервных копий"
        $dialog.ShowNewFolderButton = $true
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $TextDest.Text = $dialog.SelectedPath
        }
    })

    $BaseDialog.AcceptButton = $ButtonOK
    $BaseDialog.CancelButton = $ButtonCancel

    if ($BaseDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Валидация
        if ([string]::IsNullOrWhiteSpace($TextName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Введите имя базы!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if ([string]::IsNullOrWhiteSpace($TextSource.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Выберите путь к базе данных!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if ([string]::IsNullOrWhiteSpace($TextDest.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Выберите папку для резервных копий!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Проверка уникальности имени
        $existingBase = $Script:AllBases | Where-Object { $_.Tag -eq $TextName.Text -and ($editIndex -eq -1 -or $Script:AllBases.IndexOf($_) -ne $editIndex) }
        if ($existingBase) {
            [System.Windows.Forms.MessageBox]::Show("База с таким именем уже существует!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Создание объекта базы
        $baseConfig = [ordered]@{
            Tag = $TextName.Text.Trim()
            BackupType = $ComboType.Text
            SourcePath = $TextSource.Text.Trim()
            DestinationPath = $TextDest.Text.Trim()
            Keep = [int]$NumericKeep.Value
            CloudType = if ($CheckBoxUseCloud.Checked) { 'Yandex.Disk' } else { '' }
            CloudKeep = if ($CheckBoxUseCloud.Checked) { [int]$NumericCloudKeep.Value } else { 0 }
        }

        # Добавляем индивидуальные настройки Telegram
        if ($CheckBoxTelegramOverride.Checked) {
            $baseConfig.TelegramChatId = $TextTelegramChatId.Text.Trim()
            $baseConfig.TelegramOnlyErrors = $CheckBoxTelegramOnlyErrors.Checked
        }

        # Добавляем DT-специфичные настройки
        if ($ComboType.Text -eq "DT") {
            if (-not [string]::IsNullOrWhiteSpace($TextExe.Text)) {
                $baseConfig.ExePath = $TextExe.Text.Trim()
            }

            # Определение служб для остановки
            if ($RadioAutoServices.Checked) {
                $webServices = Get-WebServices
                if ($webServices -and ($webServices.Name -contains 'Apache2.4')) {
                    $baseConfig.StopServices = @('Apache2.4')
                } elseif ($webServices) {
                    $baseConfig.StopServices = $webServices | Select-Object -ExpandProperty Name
                } else {
                    $baseConfig.StopServices = @()
                }
            } elseif ($RadioManualServices.Checked -and $Script:SelectedServices) {
                $baseConfig.StopServices = $Script:SelectedServices
            } else {
                $baseConfig.StopServices = @()
            }
        }

        # Сохранение секретов
        if ($CheckBoxUseCloud.Checked -and -not [string]::IsNullOrWhiteSpace($TextCloudToken.Text)) {
            $tokenKey = "$($baseConfig.Tag)__YADiskToken"
            $Script:AllSecrets[$tokenKey] = $TextCloudToken.Text.Trim()
        }

        # Сохранение логина и пароля DT
        if ($ComboType.Text -eq "DT") {
            if (-not [string]::IsNullOrWhiteSpace($TextDTLogin.Text)) {
                $loginKey = "$($baseConfig.Tag)__DT_Login"
                $Script:AllSecrets[$loginKey] = $TextDTLogin.Text.Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($TextDTPassword.Text)) {
                $passwordKey = "$($baseConfig.Tag)__DT_Password"
                $Script:AllSecrets[$passwordKey] = $TextDTPassword.Text
            }
        }

        # Добавление или обновление базы
        if ($EditMode -and $editIndex -ge 0) {
            $Script:AllBases[$editIndex] = $baseConfig
        } else {
            $Script:AllBases += $baseConfig
        }

        Update-WizardBasesListView
    }

    $BaseDialog.Dispose()
}

function Show-ServicesSelectionDialog {
    $webServices = Get-WebServices
    if (-not $webServices) {
        [System.Windows.Forms.MessageBox]::Show("Веб-службы не найдены", "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return @()
    }

    $ServicesDialog = New-Object System.Windows.Forms.Form
    $ServicesDialog.Text = "Выбор служб для остановки"
    $ServicesDialog.Size = New-Object System.Drawing.Size(500, 400)
    $ServicesDialog.StartPosition = "CenterParent"
    $ServicesDialog.FormBorderStyle = "FixedDialog"
    $ServicesDialog.MaximizeBox = $false

    $LabelInfo = New-Object System.Windows.Forms.Label
    $LabelInfo.Text = "Выберите службы, которые нужно останавливать при выгрузке DT:"
    $LabelInfo.Location = New-Object System.Drawing.Point(15, 15)
    $LabelInfo.Size = New-Object System.Drawing.Size(450, 20)

    $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $CheckedListBox.Location = New-Object System.Drawing.Point(15, 45)
    $CheckedListBox.Size = New-Object System.Drawing.Size(450, 280)

    foreach ($service in $webServices) {
        $displayText = "{0} [{1}] - {2}" -f $service.DisplayName, $service.Name, $service.Status
        [void]$CheckedListBox.Items.Add($displayText)
    }

    $ButtonOK = New-Object System.Windows.Forms.Button
    $ButtonOK.Text = "OK"
    $ButtonOK.Location = New-Object System.Drawing.Point(300, 335)
    $ButtonOK.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $ButtonCancel = New-Object System.Windows.Forms.Button
    $ButtonCancel.Text = "Отмена"
    $ButtonCancel.Location = New-Object System.Drawing.Point(390, 335)
    $ButtonCancel.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $ServicesDialog.Controls.AddRange(@($LabelInfo, $CheckedListBox, $ButtonOK, $ButtonCancel))
    $ServicesDialog.AcceptButton = $ButtonOK
    $ServicesDialog.CancelButton = $ButtonCancel

    if ($ServicesDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedServices = @()
        for ($i = 0; $i -lt $CheckedListBox.CheckedItems.Count; $i++) {
            $serviceName = $webServices[$CheckedListBox.CheckedIndices[$i]].Name
            $selectedServices += $serviceName
        }
        return $selectedServices
    }

    return @()
}

function Show-WizardStep2 {
    $WizardContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Настройка уведомлений (необязательно)"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(880, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Информация
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Text = "На этом шаге вы можете настроить уведомления о результатах резервного копирования.`nЭтот шаг можно пропустить - настройка работает и без уведомлений."
    $InfoLabel.Location = New-Object System.Drawing.Point(20, 70)
    $InfoLabel.Size = New-Object System.Drawing.Size(880, 40)
    $InfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $InfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen

    $Script:WizardCheckBoxTelegramEnabled = New-Object System.Windows.Forms.CheckBox
    $Script:WizardCheckBoxTelegramEnabled.Text = "Включить отправку отчётов в Telegram"
    $Script:WizardCheckBoxTelegramEnabled.Location = New-Object System.Drawing.Point(20, 130)
    $Script:WizardCheckBoxTelegramEnabled.Size = New-Object System.Drawing.Size(400, 25)
    $Script:WizardCheckBoxTelegramEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)


    # Основной бот
    $Script:WizardGroupBoxMain = New-Object System.Windows.Forms.GroupBox
    $Script:WizardGroupBoxMain.Text = "Основной бот"
    $Script:WizardGroupBoxMain.Location = New-Object System.Drawing.Point(20, 170)
    $Script:WizardGroupBoxMain.Size = New-Object System.Drawing.Size(880, 120)
    $Script:WizardGroupBoxMain.Enabled = $false

    $LabelToken = New-Object System.Windows.Forms.Label
    $LabelToken.Text = "Токен бота:"
    $LabelToken.Location = New-Object System.Drawing.Point(15, 25)
    $LabelToken.Size = New-Object System.Drawing.Size(100, 20)

    $Script:WizardTextTelegramToken = New-Object System.Windows.Forms.TextBox
    $Script:WizardTextTelegramToken.Location = New-Object System.Drawing.Point(15, 50)
    $Script:WizardTextTelegramToken.Size = New-Object System.Drawing.Size(300, 25)
    $Script:WizardTextTelegramToken.UseSystemPasswordChar = $true


    $LabelChatId = New-Object System.Windows.Forms.Label
    $LabelChatId.Text = "ID чата:"
    $LabelChatId.Location = New-Object System.Drawing.Point(330, 25)
    $LabelChatId.Size = New-Object System.Drawing.Size(100, 20)

    $Script:WizardTextTelegramChatId = New-Object System.Windows.Forms.TextBox
    $Script:WizardTextTelegramChatId.Location = New-Object System.Drawing.Point(330, 50)
    $Script:WizardTextTelegramChatId.Size = New-Object System.Drawing.Size(150, 25)


    $Script:WizardCheckBoxOnlyErrors = New-Object System.Windows.Forms.CheckBox
    $Script:WizardCheckBoxOnlyErrors.Text = "Отправлять только при ошибках"
    $Script:WizardCheckBoxOnlyErrors.Location = New-Object System.Drawing.Point(15, 85)
    $Script:WizardCheckBoxOnlyErrors.Size = New-Object System.Drawing.Size(300, 25)

    $Script:WizardButtonTestTelegram = New-Object System.Windows.Forms.Button
    $Script:WizardButtonTestTelegram.Text = "Тест"
    $Script:WizardButtonTestTelegram.Location = New-Object System.Drawing.Point(490, 49)
    $Script:WizardButtonTestTelegram.Size = New-Object System.Drawing.Size(80, 27)
    $Script:WizardButtonTestTelegram.BackColor = [System.Drawing.Color]::LightBlue

    $Script:WizardGroupBoxMain.Controls.AddRange(@($LabelToken, $Script:WizardTextTelegramToken, $LabelChatId, $Script:WizardTextTelegramChatId, $Script:WizardCheckBoxOnlyErrors, $Script:WizardButtonTestTelegram))

    # Второй бот
    $Script:WizardGroupBoxSecondary = New-Object System.Windows.Forms.GroupBox
    $Script:WizardGroupBoxSecondary.Text = "Дополнительный бот для ответственного"
    $Script:WizardGroupBoxSecondary.Location = New-Object System.Drawing.Point(20, 310)
    $Script:WizardGroupBoxSecondary.Size = New-Object System.Drawing.Size(880, 120)
    $Script:WizardGroupBoxSecondary.Enabled = $false

    $Script:WizardCheckBoxSecondaryBot = New-Object System.Windows.Forms.CheckBox
    $Script:WizardCheckBoxSecondaryBot.Text = "Включить дополнительный бот"
    $Script:WizardCheckBoxSecondaryBot.Location = New-Object System.Drawing.Point(15, 25)
    $Script:WizardCheckBoxSecondaryBot.Size = New-Object System.Drawing.Size(300, 25)

    $LabelSecondaryToken = New-Object System.Windows.Forms.Label
    $LabelSecondaryToken.Text = "Токен второго бота:"
    $LabelSecondaryToken.Location = New-Object System.Drawing.Point(15, 55)
    $LabelSecondaryToken.Size = New-Object System.Drawing.Size(120, 20)

    $Script:WizardTextSecondaryToken = New-Object System.Windows.Forms.TextBox
    $Script:WizardTextSecondaryToken.Location = New-Object System.Drawing.Point(15, 80)
    $Script:WizardTextSecondaryToken.Size = New-Object System.Drawing.Size(300, 25)
    $Script:WizardTextSecondaryToken.UseSystemPasswordChar = $true
    $Script:WizardTextSecondaryToken.Enabled = $false

    $LabelSecondaryChatId = New-Object System.Windows.Forms.Label
    $LabelSecondaryChatId.Text = "ID чата второго бота:"
    $LabelSecondaryChatId.Location = New-Object System.Drawing.Point(330, 55)
    $LabelSecondaryChatId.Size = New-Object System.Drawing.Size(140, 20)

    $Script:WizardTextSecondaryChatId = New-Object System.Windows.Forms.TextBox
    $Script:WizardTextSecondaryChatId.Location = New-Object System.Drawing.Point(330, 80)
    $Script:WizardTextSecondaryChatId.Size = New-Object System.Drawing.Size(150, 25)
    $Script:WizardTextSecondaryChatId.Enabled = $false

    $Script:WizardButtonTestSecondary = New-Object System.Windows.Forms.Button
    $Script:WizardButtonTestSecondary.Text = "Тест"
    $Script:WizardButtonTestSecondary.Location = New-Object System.Drawing.Point(490, 79)
    $Script:WizardButtonTestSecondary.Size = New-Object System.Drawing.Size(80, 27)
    $Script:WizardButtonTestSecondary.BackColor = [System.Drawing.Color]::LightBlue
    $Script:WizardButtonTestSecondary.Enabled = $false

    $Script:WizardGroupBoxSecondary.Controls.AddRange(@($Script:WizardCheckBoxSecondaryBot, $LabelSecondaryToken, $Script:WizardTextSecondaryToken, $LabelSecondaryChatId, $Script:WizardTextSecondaryChatId, $Script:WizardButtonTestSecondary))

    # Справочная информация
    $HelpLabel = New-Object System.Windows.Forms.Label
    $HelpLabel.Text = "Для создания бота напишите @BotFather в Telegram.`nДля получения ID чата напишите @userinfobot в Telegram."
    $HelpLabel.Location = New-Object System.Drawing.Point(20, 440)
    $HelpLabel.Size = New-Object System.Drawing.Size(880, 40)
    $HelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $HelpLabel.ForeColor = [System.Drawing.Color]::Gray

    $WizardContentPanel.Controls.AddRange(@($TitleLabel, $InfoLabel, $Script:WizardCheckBoxTelegramEnabled, $Script:WizardGroupBoxMain, $Script:WizardGroupBoxSecondary, $HelpLabel))

    # Обработчики событий
    $Script:WizardCheckBoxTelegramEnabled.Add_CheckedChanged({
        $enabled = $Script:WizardCheckBoxTelegramEnabled.Checked
        $Script:WizardGroupBoxMain.Enabled = $enabled
        $Script:WizardGroupBoxSecondary.Enabled = $enabled
        if (-not $enabled -and $Script:WizardCheckBoxSecondaryBot) {
            $Script:WizardCheckBoxSecondaryBot.Checked = $false
        }
    })

    $Script:WizardCheckBoxSecondaryBot.Add_CheckedChanged({
        $enabled = $Script:WizardCheckBoxSecondaryBot.Checked
        $Script:WizardTextSecondaryToken.Enabled = $enabled
        $Script:WizardTextSecondaryChatId.Enabled = $enabled
        $Script:WizardButtonTestSecondary.Enabled = $enabled
    })

    $Script:WizardButtonTestTelegram.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($Script:WizardTextTelegramToken.Text) -or [string]::IsNullOrWhiteSpace($Script:WizardTextTelegramChatId.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Заполните токен и ID чата основного бота!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            Test-TelegramConnection -Token $Script:WizardTextTelegramToken.Text -ChatId $Script:WizardTextTelegramChatId.Text
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $Script:WizardButtonTestSecondary.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($Script:WizardTextSecondaryToken.Text) -or [string]::IsNullOrWhiteSpace($Script:WizardTextSecondaryChatId.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Заполните токен и ID чата второго бота!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            Test-TelegramConnection -Token $Script:WizardTextSecondaryToken.Text -ChatId $Script:WizardTextSecondaryChatId.Text
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Загрузка существующих настроек
    $settingsData = Load-ExistingSettings
    if ($settingsData.ContainsKey('Telegram')) {
        $telegramData = ConvertTo-Hashtable $settingsData['Telegram']
        if ($telegramData.ContainsKey('Enabled')) { $Script:WizardCheckBoxTelegramEnabled.Checked = [bool]$telegramData['Enabled'] }
        if ($telegramData.ContainsKey('ChatId')) { $Script:WizardTextTelegramChatId.Text = [string]$telegramData['ChatId'] }
        if ($telegramData.ContainsKey('NotifyOnlyOnErrors')) { $Script:WizardCheckBoxOnlyErrors.Checked = [bool]$telegramData['NotifyOnlyOnErrors'] }
        if ($telegramData.ContainsKey('SecondaryChatId') -and $telegramData['SecondaryChatId']) {
            $Script:WizardTextSecondaryChatId.Text = [string]$telegramData['SecondaryChatId']
            $Script:WizardCheckBoxSecondaryBot.Checked = $true
        }
    }

    # Загрузка токенов из секретов
    if ($Script:AllSecrets.ContainsKey('__TELEGRAM_BOT_TOKEN')) {
        $Script:WizardTextTelegramToken.Text = [string]$Script:AllSecrets['__TELEGRAM_BOT_TOKEN']
    }
    if ($Script:AllSecrets.ContainsKey('__TELEGRAM_SECONDARY_BOT_TOKEN')) {
        $Script:WizardTextSecondaryToken.Text = [string]$Script:AllSecrets['__TELEGRAM_SECONDARY_BOT_TOKEN']
    }

    # Первоначальная настройка состояния
    $enabled = $Script:WizardCheckBoxTelegramEnabled.Checked
    $Script:WizardGroupBoxMain.Enabled = $enabled
    $Script:WizardGroupBoxSecondary.Enabled = $enabled

    $secondaryEnabled = $Script:WizardCheckBoxSecondaryBot.Checked
    $Script:WizardTextSecondaryToken.Enabled = $secondaryEnabled
    $Script:WizardTextSecondaryChatId.Enabled = $secondaryEnabled
    $Script:WizardButtonTestSecondary.Enabled = $secondaryEnabled
}

function Show-WizardStep3 {
    $WizardContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Завершение настройки"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(880, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Действие после бэкапа
    $GroupBoxAfter = New-Object System.Windows.Forms.GroupBox
    $GroupBoxAfter.Text = "Действие после завершения резервного копирования"
    $GroupBoxAfter.Location = New-Object System.Drawing.Point(20, 70)
    $GroupBoxAfter.Size = New-Object System.Drawing.Size(880, 120)

    $Script:WizardRadioShutdown = New-Object System.Windows.Forms.RadioButton
    $Script:WizardRadioShutdown.Text = "Выключить компьютер"
    $Script:WizardRadioShutdown.Location = New-Object System.Drawing.Point(15, 30)
    $Script:WizardRadioShutdown.Size = New-Object System.Drawing.Size(250, 25)

    $Script:WizardRadioRestart = New-Object System.Windows.Forms.RadioButton
    $Script:WizardRadioRestart.Text = "Перезагрузить компьютер"
    $Script:WizardRadioRestart.Location = New-Object System.Drawing.Point(15, 60)
    $Script:WizardRadioRestart.Size = New-Object System.Drawing.Size(250, 25)

    $Script:WizardRadioNothing = New-Object System.Windows.Forms.RadioButton
    $Script:WizardRadioNothing.Text = "Ничего не делать"
    $Script:WizardRadioNothing.Location = New-Object System.Drawing.Point(15, 90)
    $Script:WizardRadioNothing.Size = New-Object System.Drawing.Size(250, 25)
    $Script:WizardRadioNothing.Checked = $true

    $GroupBoxAfter.Controls.AddRange(@($Script:WizardRadioShutdown, $Script:WizardRadioRestart, $Script:WizardRadioNothing))

    # Сводка настроек
    $GroupBoxSummary = New-Object System.Windows.Forms.GroupBox
    $GroupBoxSummary.Text = "Сводка настроек"
    $GroupBoxSummary.Location = New-Object System.Drawing.Point(20, 210)
    $GroupBoxSummary.Size = New-Object System.Drawing.Size(880, 250)

    $Script:WizardTextSummary = New-Object System.Windows.Forms.TextBox
    $Script:WizardTextSummary.Location = New-Object System.Drawing.Point(15, 25)
    $Script:WizardTextSummary.Size = New-Object System.Drawing.Size(850, 210)
    $Script:WizardTextSummary.Multiline = $true
    $Script:WizardTextSummary.ScrollBars = "Vertical"
    $Script:WizardTextSummary.ReadOnly = $true
    $Script:WizardTextSummary.Font = New-Object System.Drawing.Font("Consolas", 9)

    $GroupBoxSummary.Controls.Add($Script:WizardTextSummary)

    $WizardContentPanel.Controls.AddRange(@($TitleLabel, $GroupBoxAfter, $GroupBoxSummary))

    # Загрузка существующих настроек
    $settingsData = Load-ExistingSettings
    if ($settingsData.ContainsKey('AfterBackup')) {
        switch ([int]$settingsData['AfterBackup']) {
            1 { $Script:WizardRadioShutdown.Checked = $true }
            2 { $Script:WizardRadioRestart.Checked = $true }
            default { $Script:WizardRadioNothing.Checked = $true }
        }
    }

    Update-WizardSummary
}

function Update-WizardSummary {
    $summary = @()
    $summary += "=== НАСТРОЙКИ СИСТЕМЫ РЕЗЕРВНОГО КОПИРОВАНИЯ 1С ==="
    $summary += ""

    # Базы данных
    $summary += "БАЗЫ ДАННЫХ ($($Script:AllBases.Count)):"
    if ($Script:AllBases.Count -eq 0) {
        $summary += "  Не настроено ни одной базы"
    } else {
        foreach ($base in $Script:AllBases) {
            $summary += "  • $($base.Tag) [$($base.BackupType)]"
            $summary += "    Источник: $($base.SourcePath)"
            $summary += "    Назначение: $($base.DestinationPath)"
            $summary += "    Копий: $($base.Keep)"
            $summary += ""
        }
    }

    # Telegram
    $summary += "УВЕДОМЛЕНИЯ TELEGRAM:"
    if ($Script:WizardCheckBoxTelegramEnabled -and $Script:WizardCheckBoxTelegramEnabled.Checked) {
        $summary += "  Включено"
        $summary += "  Chat ID: $($Script:WizardTextTelegramChatId.Text)"
        $summary += "  Только ошибки: $(if ($Script:WizardCheckBoxOnlyErrors.Checked) { 'Да' } else { 'Нет' })"
    } else {
        $summary += "  Отключено"
    }
    $summary += ""

    # Действие после бэкапа
    $summary += "ДЕЙСТВИЕ ПОСЛЕ РЕЗЕРВНОГО КОПИРОВАНИЯ:"
    if ($Script:WizardRadioShutdown -and $Script:WizardRadioShutdown.Checked) {
        $summary += "  Выключение компьютера"
    } elseif ($Script:WizardRadioRestart -and $Script:WizardRadioRestart.Checked) {
        $summary += "  Перезагрузка компьютера"
    } else {
        $summary += "  Без действий"
    }

    if ($Script:WizardTextSummary) {
        $Script:WizardTextSummary.Text = $summary -join "`r`n"
    }
}

function Update-WizardProgress {
    $WizardProgressBar.Value = $Script:CurrentStep
    switch ($Script:CurrentStep) {
        1 { $WizardProgressLabel.Text = "Шаг 1 из 3: Настройка баз данных" }
        2 { $WizardProgressLabel.Text = "Шаг 2 из 3: Настройка уведомлений" }
        3 { $WizardProgressLabel.Text = "Шаг 3 из 3: Завершение настройки" }
    }

    $WizardButtonBack.Enabled = ($Script:CurrentStep -gt 1)
    $WizardButtonNext.Visible = ($Script:CurrentStep -lt 3)
    $WizardButtonFinish.Visible = ($Script:CurrentStep -eq 3)

    switch ($Script:CurrentStep) {
        1 { Show-WizardStep1 }
        2 { Show-WizardStep2 }
        3 { Show-WizardStep3 }
    }
}

function Save-WizardSettings {
    # Сохранение конфигураций баз
    foreach ($base in $Script:AllBases) {
        $configPath = Join-Path $BasesDir "$($base.Tag).json"
        $base | ConvertTo-Json -Depth 8 | Set-Content -Path $configPath -Encoding UTF8
    }

    # Подготовка секретов
    if ($Script:WizardCheckBoxTelegramEnabled.Checked -and -not [string]::IsNullOrWhiteSpace($Script:WizardTextTelegramToken.Text)) {
        $Script:AllSecrets['__TELEGRAM_BOT_TOKEN'] = $Script:WizardTextTelegramToken.Text
    } elseif ($Script:AllSecrets.ContainsKey('__TELEGRAM_BOT_TOKEN')) {
        $Script:AllSecrets.Remove('__TELEGRAM_BOT_TOKEN')
    }

    if ($Script:WizardCheckBoxSecondaryBot.Checked -and -not [string]::IsNullOrWhiteSpace($Script:WizardTextSecondaryToken.Text)) {
        $Script:AllSecrets['__TELEGRAM_SECONDARY_BOT_TOKEN'] = $Script:WizardTextSecondaryToken.Text
    } elseif ($Script:AllSecrets.ContainsKey('__TELEGRAM_SECONDARY_BOT_TOKEN')) {
        $Script:AllSecrets.Remove('__TELEGRAM_SECONDARY_BOT_TOKEN')
    }

    # Сохранение секретов
    Encrypt-Secrets -Secrets $Script:AllSecrets -KeyPath $KeyPath -OutFile $SecretsFile

    # Настройки Telegram
    $telegramSettings = @{
        Enabled = if ($Script:WizardCheckBoxTelegramEnabled) { $Script:WizardCheckBoxTelegramEnabled.Checked } else { $false }
        ChatId = if ($Script:WizardTextTelegramChatId) { $Script:WizardTextTelegramChatId.Text.Trim() } else { '' }
        NotifyOnlyOnErrors = if ($Script:WizardCheckBoxOnlyErrors) { $Script:WizardCheckBoxOnlyErrors.Checked } else { $false }
        SecondaryChatId = if ($Script:WizardCheckBoxSecondaryBot.Checked -and $Script:WizardTextSecondaryChatId) { $Script:WizardTextSecondaryChatId.Text.Trim() } else { '' }
    }

    # Действие после бэкапа
    $afterBackup = 3
    if ($Script:WizardRadioShutdown.Checked) { $afterBackup = 1 }
    elseif ($Script:WizardRadioRestart.Checked) { $afterBackup = 2 }

    # Общие настройки
    $settings = @{
        AfterBackup = $afterBackup
        Telegram = $telegramSettings
    }

    $settings | ConvertTo-Json -Depth 6 | Set-Content -Path $SettingsFile -Encoding UTF8
}

# Обработчики навигации мастера
$WizardButtonNext.Add_Click({
    if ($Script:CurrentStep -eq 1) {
        if ($Script:AllBases.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Настройте хотя бы одну базу данных!", "Внимание", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }

    if ($Script:CurrentStep -lt $Script:MaxSteps) {
        $Script:CurrentStep++
        Update-WizardProgress
    }
})

$WizardButtonBack.Add_Click({
    if ($Script:CurrentStep -gt 1) {
        $Script:CurrentStep--
        Update-WizardProgress
    }
})

$WizardButtonFinish.Add_Click({
    try {
        Save-WizardSettings
        [System.Windows.Forms.MessageBox]::Show("Настройка успешно завершена!`n`nТеперь вы можете запускать резервное копирование через Run-Backup.cmd", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # Переключиться на вкладку редактора и обновить список баз
        $TabControl.SelectedTab = $TabEditor
        Load-EditorBasesList
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при сохранении настроек:`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ==========================================
# РЕДАКТОР КОНФИГУРАЦИЙ
# ==========================================

# === ЛЕВАЯ ПАНЕЛЬ СО СПИСКОМ БАЗ ===
$EditorGroupBoxList = New-Object System.Windows.Forms.GroupBox
$EditorGroupBoxList.Text = "Базы данных"
$EditorGroupBoxList.Location = New-Object System.Drawing.Point(12, 12)
$EditorGroupBoxList.Size = New-Object System.Drawing.Size(250, 850)
$EditorGroupBoxList.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$EditorListBoxBases = New-Object System.Windows.Forms.ListBox
$EditorListBoxBases.Location = New-Object System.Drawing.Point(10, 25)
$EditorListBoxBases.Size = New-Object System.Drawing.Size(230, 520)
$EditorListBoxBases.Font = New-Object System.Drawing.Font("Consolas", 10)
$EditorListBoxBases.IntegralHeight = $false

$EditorButtonRefresh = New-Object System.Windows.Forms.Button
$EditorButtonRefresh.Text = "Обновить"
$EditorButtonRefresh.Location = New-Object System.Drawing.Point(10, 555)
$EditorButtonRefresh.Size = New-Object System.Drawing.Size(110, 35)
$EditorButtonRefresh.UseVisualStyleBackColor = $true

$EditorButtonDelete = New-Object System.Windows.Forms.Button
$EditorButtonDelete.Text = "Удалить"
$EditorButtonDelete.Location = New-Object System.Drawing.Point(130, 555)
$EditorButtonDelete.Size = New-Object System.Drawing.Size(110, 35)
$EditorButtonDelete.UseVisualStyleBackColor = $true
$EditorButtonDelete.Enabled = $false

$EditorGroupBoxList.Controls.AddRange(@($EditorListBoxBases, $EditorButtonRefresh, $EditorButtonDelete))

# === ПРАВАЯ ПАНЕЛЬ С ПАРАМЕТРАМИ ===
$EditorGroupBoxParams = New-Object System.Windows.Forms.GroupBox
$EditorGroupBoxParams.Text = "Параметры базы"
$EditorGroupBoxParams.Location = New-Object System.Drawing.Point(275, 12)
$EditorGroupBoxParams.Size = New-Object System.Drawing.Size(865, 720)
$EditorGroupBoxParams.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$EditorGroupBoxParams.Enabled = $false

# Тип бэкапа
$EditorLabelType = New-Object System.Windows.Forms.Label
$EditorLabelType.Text = "Тип резервной копии:"
$EditorLabelType.Location = New-Object System.Drawing.Point(15, 30)
$EditorLabelType.Size = New-Object System.Drawing.Size(200, 20)


$EditorComboType = New-Object System.Windows.Forms.ComboBox
$EditorComboType.Items.AddRange(@("1CD", "DT"))
$EditorComboType.DropDownStyle = "DropDownList"
$EditorComboType.Location = New-Object System.Drawing.Point(15, 55)
$EditorComboType.Size = New-Object System.Drawing.Size(620, 25)

# Путь к базе
$EditorLabelSource = New-Object System.Windows.Forms.Label
$EditorLabelSource.Text = "Путь к базе данных:"
$EditorLabelSource.Location = New-Object System.Drawing.Point(15, 95)
$EditorLabelSource.Size = New-Object System.Drawing.Size(200, 20)


$EditorTextSource = New-Object System.Windows.Forms.TextBox
$EditorTextSource.Location = New-Object System.Drawing.Point(15, 120)
$EditorTextSource.Size = New-Object System.Drawing.Size(530, 25)

$EditorButtonBrowseSource = New-Object System.Windows.Forms.Button
$EditorButtonBrowseSource.Text = "Обзор..."
$EditorButtonBrowseSource.Location = New-Object System.Drawing.Point(555, 119)
$EditorButtonBrowseSource.Size = New-Object System.Drawing.Size(80, 27)
$EditorButtonBrowseSource.UseVisualStyleBackColor = $true

# Папка для бэкапов
$EditorLabelDest = New-Object System.Windows.Forms.Label
$EditorLabelDest.Text = "Папка для резервных копий:"
$EditorLabelDest.Location = New-Object System.Drawing.Point(15, 160)
$EditorLabelDest.Size = New-Object System.Drawing.Size(250, 20)

$EditorTextDest = New-Object System.Windows.Forms.TextBox
$EditorTextDest.Location = New-Object System.Drawing.Point(15, 185)
$EditorTextDest.Size = New-Object System.Drawing.Size(530, 25)

$EditorButtonBrowseDest = New-Object System.Windows.Forms.Button
$EditorButtonBrowseDest.Text = "Обзор..."
$EditorButtonBrowseDest.Location = New-Object System.Drawing.Point(555, 184)
$EditorButtonBrowseDest.Size = New-Object System.Drawing.Size(80, 27)
$EditorButtonBrowseDest.UseVisualStyleBackColor = $true

# Количество копий
$EditorLabelKeep = New-Object System.Windows.Forms.Label
$EditorLabelKeep.Text = "Хранить копий локально:"
$EditorLabelKeep.Location = New-Object System.Drawing.Point(15, 225)
$EditorLabelKeep.Size = New-Object System.Drawing.Size(170, 20)

$EditorNumericKeep = New-Object System.Windows.Forms.NumericUpDown
$EditorNumericKeep.Location = New-Object System.Drawing.Point(15, 250)
$EditorNumericKeep.Size = New-Object System.Drawing.Size(100, 25)
$EditorNumericKeep.Minimum = 1
$EditorNumericKeep.Maximum = 999
$EditorNumericKeep.Value = 3

# Облачное хранилище
$EditorLabelCloud = New-Object System.Windows.Forms.Label
$EditorLabelCloud.Text = "Облачное хранилище:"
$EditorLabelCloud.Location = New-Object System.Drawing.Point(200, 225)
$EditorLabelCloud.Size = New-Object System.Drawing.Size(180, 20)

$EditorComboCloud = New-Object System.Windows.Forms.ComboBox
$EditorComboCloud.Items.AddRange(@("Не использовать", "Yandex.Disk"))
$EditorComboCloud.DropDownStyle = "DropDownList"
$EditorComboCloud.Location = New-Object System.Drawing.Point(200, 250)
$EditorComboCloud.Size = New-Object System.Drawing.Size(180, 25)

$EditorLabelCloudKeep = New-Object System.Windows.Forms.Label
$EditorLabelCloudKeep.Text = "Копий в облаке:"
$EditorLabelCloudKeep.Location = New-Object System.Drawing.Point(400, 225)
$EditorLabelCloudKeep.Size = New-Object System.Drawing.Size(120, 20)

$EditorNumericCloudKeep = New-Object System.Windows.Forms.NumericUpDown
$EditorNumericCloudKeep.Location = New-Object System.Drawing.Point(400, 250)
$EditorNumericCloudKeep.Size = New-Object System.Drawing.Size(100, 25)
$EditorNumericCloudKeep.Minimum = 0
$EditorNumericCloudKeep.Maximum = 999
$EditorNumericCloudKeep.Value = 0

# DT Настройки группа
$EditorGroupBoxDT = New-Object System.Windows.Forms.GroupBox
$EditorGroupBoxDT.Text = "Настройки для выгрузки DT"
$EditorGroupBoxDT.Location = New-Object System.Drawing.Point(15, 290)
$EditorGroupBoxDT.Size = New-Object System.Drawing.Size(835, 210)
$EditorGroupBoxDT.Enabled = $false

# Путь к 1cestart.exe
$EditorLabelExe = New-Object System.Windows.Forms.Label
$EditorLabelExe.Text = "Путь к 1cestart.exe (обязательно для DT):"
$EditorLabelExe.Location = New-Object System.Drawing.Point(10, 25)
$EditorLabelExe.Size = New-Object System.Drawing.Size(250, 20)
$EditorLabelExe.ForeColor = [System.Drawing.Color]::DarkRed

$EditorTextExe = New-Object System.Windows.Forms.TextBox
$EditorTextExe.Location = New-Object System.Drawing.Point(10, 50)
$EditorTextExe.Size = New-Object System.Drawing.Size(500, 25)

$EditorButtonBrowseExe = New-Object System.Windows.Forms.Button
$EditorButtonBrowseExe.Text = "Обзор..."
$EditorButtonBrowseExe.Location = New-Object System.Drawing.Point(520, 49)
$EditorButtonBrowseExe.Size = New-Object System.Drawing.Size(90, 27)
$EditorButtonBrowseExe.UseVisualStyleBackColor = $true

# Логин и пароль для DT
$EditorLabelDTLogin = New-Object System.Windows.Forms.Label
$EditorLabelDTLogin.Text = "Логин 1С:"
$EditorLabelDTLogin.Location = New-Object System.Drawing.Point(10, 85)
$EditorLabelDTLogin.Size = New-Object System.Drawing.Size(100, 20)

$EditorTextDTLogin = New-Object System.Windows.Forms.TextBox
$EditorTextDTLogin.Location = New-Object System.Drawing.Point(10, 110)
$EditorTextDTLogin.Size = New-Object System.Drawing.Size(200, 25)

$EditorLabelDTPassword = New-Object System.Windows.Forms.Label
$EditorLabelDTPassword.Text = "Пароль 1С:"
$EditorLabelDTPassword.Location = New-Object System.Drawing.Point(220, 85)
$EditorLabelDTPassword.Size = New-Object System.Drawing.Size(100, 20)

$EditorTextDTPassword = New-Object System.Windows.Forms.TextBox
$EditorTextDTPassword.Location = New-Object System.Drawing.Point(220, 110)
$EditorTextDTPassword.Size = New-Object System.Drawing.Size(200, 25)
$EditorTextDTPassword.UseSystemPasswordChar = $true

# Службы для остановки
$EditorLabelServices = New-Object System.Windows.Forms.Label
$EditorLabelServices.Text = "Веб-службы при экспорте DT:"
$EditorLabelServices.Location = New-Object System.Drawing.Point(10, 145)
$EditorLabelServices.Size = New-Object System.Drawing.Size(200, 20)

$EditorRadioAutoServices = New-Object System.Windows.Forms.RadioButton
$EditorRadioAutoServices.Text = "Автоматически (все веб-службы)"
$EditorRadioAutoServices.Location = New-Object System.Drawing.Point(10, 170)
$EditorRadioAutoServices.Size = New-Object System.Drawing.Size(250, 25)
$EditorRadioAutoServices.Checked = $true

$EditorRadioManualServices = New-Object System.Windows.Forms.RadioButton
$EditorRadioManualServices.Text = "Только Apache2.4"
$EditorRadioManualServices.Location = New-Object System.Drawing.Point(270, 170)
$EditorRadioManualServices.Size = New-Object System.Drawing.Size(150, 25)

$EditorRadioNoServices = New-Object System.Windows.Forms.RadioButton
$EditorRadioNoServices.Text = "Не останавливать"
$EditorRadioNoServices.Location = New-Object System.Drawing.Point(430, 170)
$EditorRadioNoServices.Size = New-Object System.Drawing.Size(150, 25)

$EditorGroupBoxDT.Controls.AddRange(@($EditorLabelExe, $EditorTextExe, $EditorButtonBrowseExe, $EditorLabelDTLogin, $EditorTextDTLogin, $EditorLabelDTPassword, $EditorTextDTPassword, $EditorLabelServices, $EditorRadioAutoServices, $EditorRadioManualServices, $EditorRadioNoServices))

# Облачное хранилище группа
$EditorGroupBoxCloud = New-Object System.Windows.Forms.GroupBox
$EditorGroupBoxCloud.Text = "Яндекс.Диск"
$EditorGroupBoxCloud.Location = New-Object System.Drawing.Point(15, 510)
$EditorGroupBoxCloud.Size = New-Object System.Drawing.Size(835, 100)

$EditorCheckBoxUseCloud = New-Object System.Windows.Forms.CheckBox
$EditorCheckBoxUseCloud.Text = "Отправлять копии в Яндекс.Диск"
$EditorCheckBoxUseCloud.Location = New-Object System.Drawing.Point(10, 25)
$EditorCheckBoxUseCloud.Size = New-Object System.Drawing.Size(250, 20)

$EditorLabelCloudToken = New-Object System.Windows.Forms.Label
$EditorLabelCloudToken.Text = "OAuth токен:"
$EditorLabelCloudToken.Location = New-Object System.Drawing.Point(10, 50)
$EditorLabelCloudToken.Size = New-Object System.Drawing.Size(100, 20)
$EditorLabelCloudToken.Enabled = $false

$EditorTextCloudToken = New-Object System.Windows.Forms.TextBox
$EditorTextCloudToken.Location = New-Object System.Drawing.Point(10, 75)
$EditorTextCloudToken.Size = New-Object System.Drawing.Size(300, 25)
$EditorTextCloudToken.UseSystemPasswordChar = $true
$EditorTextCloudToken.Enabled = $false

$EditorButtonTestCloud = New-Object System.Windows.Forms.Button
$EditorButtonTestCloud.Text = "Тест"
$EditorButtonTestCloud.Location = New-Object System.Drawing.Point(570, 74)
$EditorButtonTestCloud.Size = New-Object System.Drawing.Size(60, 27)
$EditorButtonTestCloud.BackColor = [System.Drawing.Color]::LightBlue
$EditorButtonTestCloud.Enabled = $false

$EditorGroupBoxCloud.Controls.AddRange(@($EditorCheckBoxUseCloud, $EditorLabelCloudToken, $EditorTextCloudToken, $EditorLabelCloudKeep, $EditorNumericCloudKeep, $EditorButtonTestCloud))

# Включение/выключение базы
$EditorCheckEnabled = New-Object System.Windows.Forms.CheckBox
$EditorCheckEnabled.Text = "База включена для резервного копирования"
$EditorCheckEnabled.Location = New-Object System.Drawing.Point(15, 630)
$EditorCheckEnabled.Size = New-Object System.Drawing.Size(400, 25)
$EditorCheckEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$EditorCheckEnabled.Checked = $true

# Кнопка сохранения
$EditorButtonSave = New-Object System.Windows.Forms.Button
$EditorButtonSave.Text = "Сохранить изменения"
$EditorButtonSave.Location = New-Object System.Drawing.Point(15, 660)
$EditorButtonSave.Size = New-Object System.Drawing.Size(200, 40)
$EditorButtonSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$EditorButtonSave.BackColor = [System.Drawing.Color]::LightGreen
$EditorButtonSave.UseVisualStyleBackColor = $false
$EditorButtonSave.Enabled = $false

$EditorGroupBoxParams.Controls.AddRange(@(
    $EditorLabelType, $EditorComboType,
    $EditorLabelSource, $EditorTextSource, $EditorButtonBrowseSource,
    $EditorLabelDest, $EditorTextDest, $EditorButtonBrowseDest,
    $EditorLabelKeep, $EditorNumericKeep,
    $EditorLabelCloud, $EditorComboCloud,
    $EditorLabelCloudKeep, $EditorNumericCloudKeep,
    $EditorGroupBoxCloud,
    $EditorGroupBoxDT,
    $EditorCheckEnabled,
    $EditorButtonSave
))

$TabEditor.Controls.AddRange(@($EditorGroupBoxList, $EditorGroupBoxParams))

# === ФУНКЦИИ РЕДАКТОРА ===
function Load-EditorBasesList {
    $EditorListBoxBases.Items.Clear()

    if (-not (Test-Path $BasesDir)) {
        $EditorListBoxBases.Items.Add("(Нет баз)")
        return
    }

    $bases = Get-ChildItem -Path $BasesDir -Filter "*.json" -File -ErrorAction SilentlyContinue

    if ($bases.Count -eq 0) {
        $EditorListBoxBases.Items.Add("(Нет баз)")
        return
    }

    foreach ($baseFile in $bases | Sort-Object Name) {
        $baseName = $baseFile.BaseName
        try {
            $config = Get-Content $baseFile.FullName -Raw | ConvertFrom-Json
            $status = if ($config.Disabled) { "[OFF]" } else { "[ON] " }
            $EditorListBoxBases.Items.Add("$status $baseName")
        } catch {
            $EditorListBoxBases.Items.Add("[ERR] $baseName")
        }
    }

    if ($EditorListBoxBases.Items.Count -gt 0 -and $EditorListBoxBases.Items[0] -ne "(Нет баз)") {
        $EditorListBoxBases.SelectedIndex = 0
    }
}

function Load-SchedulerDatabasesList {
    $SchedulerComboDatabase.Items.Clear()
    [void]$SchedulerComboDatabase.Items.Add("Все базы")

    if (-not (Test-Path $BasesDir)) {
        $SchedulerComboDatabase.SelectedIndex = 0
        return
    }

    $bases = Get-ChildItem -Path $BasesDir -Filter "*.json" -File -ErrorAction SilentlyContinue

    foreach ($baseFile in $bases | Sort-Object Name) {
        $baseName = $baseFile.BaseName
        try {
            $config = Get-Content $baseFile.FullName -Raw | ConvertFrom-Json
            if (-not $config.Disabled) {
                [void]$SchedulerComboDatabase.Items.Add($baseName)
            }
        } catch {
            # Skip invalid config files
        }
    }

    $SchedulerComboDatabase.SelectedIndex = 0
}

function Load-EditorBaseConfig {
    param([string]$BaseName)

    $Script:CurrentBase = $BaseName
    $configPath = Join-Path $BasesDir "$BaseName.json"

    if (-not (Test-Path $configPath)) { return }

    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json

        # Загружаем значения
        $EditorComboType.Text = if ($config.BackupType) { $config.BackupType } else { "1CD" }
        $EditorTextSource.Text = if ($config.SourcePath) { $config.SourcePath } else { "" }
        $EditorTextDest.Text = if ($config.DestinationPath) { $config.DestinationPath } else { "" }
        $EditorNumericKeep.Value = if ($config.Keep) { [Math]::Max(1, [int]$config.Keep) } else { 3 }

        # Облачные настройки
        $EditorComboCloud.Text = if ($config.CloudType -eq "Yandex.Disk") { "Yandex.Disk" } else { "Не использовать" }
        $EditorCheckBoxUseCloud.Checked = ($config.CloudType -eq "Yandex.Disk")

        if ($config.PSObject.Properties.Name -contains 'CloudKeep') {
            $EditorNumericCloudKeep.Value = [Math]::Max(0, [int]$config.CloudKeep)
        } else {
            $EditorNumericCloudKeep.Value = 0
        }

        # DT настройки - радио кнопки сервисов
        if ($config.PSObject.Properties.Name -contains 'StopServices' -and $config.StopServices) {
            $services = $config.StopServices
            if ($services.Count -eq 0) {
                $EditorRadioNoServices.Checked = $true
            } elseif (($services | Where-Object { $_ -eq "Apache2.4" }).Count -eq $services.Count) {
                $EditorRadioManualServices.Checked = $true
            } else {
                $EditorRadioAutoServices.Checked = $true
            }
        } else {
            $EditorRadioAutoServices.Checked = $true
        }

        if ($config.PSObject.Properties.Name -contains 'ExePath') {
            $EditorTextExe.Text = $config.ExePath
        } else {
            $EditorTextExe.Text = ""
        }

        # Загрузка секретов
        $tokenKey = "$BaseName" + "__YADiskToken"
        if ($Script:AllSecrets.ContainsKey($tokenKey)) {
            $EditorTextCloudToken.Text = [string]$Script:AllSecrets[$tokenKey]
        } else {
            $EditorTextCloudToken.Text = ""
        }

        $loginKey = "$BaseName" + "__DT_Login"
        if ($Script:AllSecrets.ContainsKey($loginKey)) {
            $EditorTextDTLogin.Text = [string]$Script:AllSecrets[$loginKey]
        } else {
            $EditorTextDTLogin.Text = ""
        }

        $passwordKey = "$BaseName" + "__DT_Password"
        if ($Script:AllSecrets.ContainsKey($passwordKey)) {
            $EditorTextDTPassword.Text = [string]$Script:AllSecrets[$passwordKey]
        } else {
            $EditorTextDTPassword.Text = ""
        }

        $EditorCheckEnabled.Checked = -not $config.Disabled

        # Включаем/отключаем поля в зависимости от типа
        $isDT = ($EditorComboType.Text -eq "DT")
        $EditorGroupBoxDT.Enabled = $isDT

        $useCloud = $EditorCheckBoxUseCloud.Checked
        $EditorLabelCloudToken.Enabled = $useCloud
        $EditorTextCloudToken.Enabled = $useCloud
        $EditorButtonTestCloud.Enabled = $useCloud

        $EditorGroupBoxParams.Enabled = $true
        $EditorButtonSave.Enabled = $false
        $EditorButtonDelete.Enabled = $true
        $Script:HasChanges = $false

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Ошибка загрузки конфигурации: $_",
            "Ошибка",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Save-EditorBaseConfig {
    if (-not $Script:CurrentBase) { return }

    $configPath = Join-Path $BasesDir "$($Script:CurrentBase).json"

    $config = [ordered]@{
        Tag = $Script:CurrentBase
        BackupType = $EditorComboType.Text
        SourcePath = $EditorTextSource.Text
        DestinationPath = $EditorTextDest.Text
        Keep = [int]$EditorNumericKeep.Value
        CloudType = if ($EditorCheckBoxUseCloud.Checked) { "Yandex.Disk" } else { "" }
        CloudKeep = if ($EditorCheckBoxUseCloud.Checked) { [int]$EditorNumericCloudKeep.Value } else { 0 }
        Disabled = -not $EditorCheckEnabled.Checked
    }

    if ($EditorComboType.Text -eq "DT") {
        # Определяем службы по выбранной радио-кнопке
        if ($EditorRadioNoServices.Checked) {
            $config.StopServices = @()
        } elseif ($EditorRadioManualServices.Checked) {
            $config.StopServices = @("Apache2.4")
        } else {
            # Автоматически - получаем список всех веб-служб
            $config.StopServices = @("Apache2.4", "httpd", "nginx", "iis", "w3svc")
        }

        if ($EditorTextExe.Text.Trim()) {
            $config.ExePath = $EditorTextExe.Text.Trim()
        }
    }

    # Сохранение секретов
    if ($EditorCheckBoxUseCloud.Checked -and -not [string]::IsNullOrWhiteSpace($EditorTextCloudToken.Text)) {
        $tokenKey = "$($Script:CurrentBase)__YADiskToken"
        $Script:AllSecrets[$tokenKey] = $EditorTextCloudToken.Text.Trim()
    }

    if ($EditorComboType.Text -eq "DT") {
        if (-not [string]::IsNullOrWhiteSpace($EditorTextDTLogin.Text)) {
            $loginKey = "$($Script:CurrentBase)__DT_Login"
            $Script:AllSecrets[$loginKey] = $EditorTextDTLogin.Text.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($EditorTextDTPassword.Text)) {
            $passwordKey = "$($Script:CurrentBase)__DT_Password"
            $Script:AllSecrets[$passwordKey] = $EditorTextDTPassword.Text
        }
    }

    try {
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "Конфигурация базы '$($Script:CurrentBase)' сохранена.",
            "Успех",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        $EditorButtonSave.Enabled = $false
        $Script:HasChanges = $false

        Load-EditorBasesList

        # Восстанавливаем выбор
        for ($i = 0; $i -lt $EditorListBoxBases.Items.Count; $i++) {
            if ($EditorListBoxBases.Items[$i] -match "$($Script:CurrentBase)$") {
                $EditorListBoxBases.SelectedIndex = $i
                break
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Ошибка сохранения: $_",
            "Ошибка",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# === ОБРАБОТЧИКИ СОБЫТИЙ РЕДАКТОРА ===
$EditorListBoxBases.Add_SelectedIndexChanged({
    if ($EditorListBoxBases.SelectedItem -and $EditorListBoxBases.SelectedItem -ne "(Нет баз)") {
        $baseName = $EditorListBoxBases.SelectedItem -replace '^\[.*?\]\s*', ''
        Load-EditorBaseConfig -BaseName $baseName
    }
})

$EditorButtonRefresh.Add_Click({ Load-EditorBasesList })

$EditorButtonDelete.Add_Click({
    if (-not $Script:CurrentBase) { return }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Удалить конфигурацию базы '$($Script:CurrentBase)'?`n`nЭто действие нельзя отменить!",
        "Подтверждение удаления",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $configPath = Join-Path $BasesDir "$($Script:CurrentBase).json"

        try {
            Remove-Item $configPath -Force
            $EditorGroupBoxParams.Enabled = $false
            $EditorButtonSave.Enabled = $false
            $EditorButtonDelete.Enabled = $false
            $Script:CurrentBase = $null
            Load-EditorBasesList
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Ошибка удаления: $_",
                "Ошибка",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

$EditorButtonSave.Add_Click({ Save-EditorBaseConfig })

# Обработчики изменений
$Script:EditorEnableSaveHandler = {
    if ($EditorGroupBoxParams.Enabled) {
        $EditorButtonSave.Enabled = $true
        $Script:HasChanges = $true
    }
}

$EditorComboType.Add_SelectedIndexChanged({
    & $Script:EditorEnableSaveHandler
    $isDT = ($EditorComboType.Text -eq "DT")
    $EditorGroupBoxDT.Enabled = $isDT
})

$EditorTextSource.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorTextDest.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorNumericKeep.Add_ValueChanged($Script:EditorEnableSaveHandler)
$EditorComboCloud.Add_SelectedIndexChanged($Script:EditorEnableSaveHandler)
$EditorNumericCloudKeep.Add_ValueChanged($Script:EditorEnableSaveHandler)
$EditorTextExe.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorRadioAutoServices.Add_CheckedChanged($Script:EditorEnableSaveHandler)
$EditorRadioManualServices.Add_CheckedChanged($Script:EditorEnableSaveHandler)
$EditorRadioNoServices.Add_CheckedChanged($Script:EditorEnableSaveHandler)
$EditorTextDTLogin.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorTextDTPassword.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorCheckBoxUseCloud.Add_CheckedChanged($Script:EditorEnableSaveHandler)
$EditorTextCloudToken.Add_TextChanged($Script:EditorEnableSaveHandler)
$EditorCheckEnabled.Add_CheckedChanged($Script:EditorEnableSaveHandler)

# Обработчик изменения облачных настроек
$EditorCheckBoxUseCloud.Add_CheckedChanged({
    & $Script:EditorEnableSaveHandler
    $useCloud = $EditorCheckBoxUseCloud.Checked
    $EditorLabelCloudToken.Enabled = $useCloud
    $EditorTextCloudToken.Enabled = $useCloud
    $EditorButtonTestCloud.Enabled = $useCloud
})

# Обработчики кнопок тестирования
$EditorButtonTestCloud.Add_Click({
    if ([string]::IsNullOrWhiteSpace($EditorTextCloudToken.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Введите токен Яндекс.Диска!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $baseName = if ($Script:CurrentBase) { $Script:CurrentBase } else { "Test" }
    Test-YandexDiskConnection -Token $EditorTextCloudToken.Text -BaseName $baseName
})

# Обработчики кнопок обзора
$EditorButtonBrowseSource.Add_Click({
    if ($EditorComboType.Text -eq "1CD") {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "База 1С (*.1CD)|*.1CD|Все файлы (*.*)|*.*"
        $dialog.Title = "Выберите файл базы 1С"

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $EditorTextSource.Text = $dialog.FileName
        }
    } else {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Выберите папку с базой 1С"
        $dialog.ShowNewFolderButton = $false

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $EditorTextSource.Text = $dialog.SelectedPath
        }
    }
})

$EditorButtonBrowseDest.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Выберите папку для резервных копий"
    $dialog.ShowNewFolderButton = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $EditorTextDest.Text = $dialog.SelectedPath
    }
})

$EditorButtonBrowseExe.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "1cestart.exe|1cestart.exe|Все exe файлы (*.exe)|*.exe"
    $dialog.Title = "Выберите 1cestart.exe"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $EditorTextExe.Text = $dialog.FileName
    }
})

# === ОБРАБОТЧИКИ ОБЩИХ СОБЫТИЙ ===
$TabControl.Add_SelectedIndexChanged({
    if ($TabControl.SelectedTab -eq $TabEditor) {
        Load-EditorBasesList
    } elseif ($TabControl.SelectedTab -eq $TabScheduler) {
        Load-SchedulerDatabasesList
    }
})

# Обработчик закрытия
$MainForm.Add_FormClosing({
    param($sender, $e)

    if ($Script:HasChanges) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Есть несохраненные изменения. Сохранить перед закрытием?",
            "Несохраненные изменения",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        switch ($result) {
            "Yes" { Save-EditorBaseConfig }
            "Cancel" { $e.Cancel = $true }
        }
    }
})

# === ФУНКЦИИ ТЕСТИРОВАНИЯ ===
function Test-YandexDiskConnection {
    param(
        [Parameter(Mandatory)][string]$Token,
        [string]$BaseName = 'Test'
    )

    try {
        # Проверяем наличие модуля
        $yandexModulePath = Join-Path $ProjectRoot 'modules\Cloud.YandexDisk.psm1'
        if (-not (Test-Path $yandexModulePath)) {
            [System.Windows.Forms.MessageBox]::Show("Модуль Cloud.YandexDisk.psm1 не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        Import-Module -Force -DisableNameChecking $yandexModulePath -ErrorAction Stop

        # Создаём тестовый файл
        $tmpFile = [IO.Path]::GetTempFileName()
        $testContent = "Тест подключения к Яндекс.Диск `nДата и время: {0}`nОт системы резервного копирования 1С" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        [IO.File]::WriteAllText($tmpFile, $testContent, [System.Text.Encoding]::UTF8)

        try {
            # Проверяем доступ к диску
            $remoteFolder = "/Backups1C/$BaseName"
            $remoteTest = "$remoteFolder/__test_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date)

            # Создаём папку если нужно
            Ensure-YandexDiskFolder -Token $Token -RemotePath $remoteFolder

            # Загружаем тестовый файл
            Upload-ToYandexDisk -Token $Token -LocalPath $tmpFile -RemotePath $remoteTest -BarWidth 10

            # Удаляем тестовый файл
            try {
                $headers = @{ Authorization = "OAuth $Token" }
                $encPath = [Uri]::EscapeDataString($remoteTest)
                Invoke-RestMethod -Uri "https://cloud-api.yandex.net/v1/disk/resources?path=$encPath&permanently=true" -Headers $headers -Method Delete -ErrorAction Stop | Out-Null
            }
            catch {
                # Не критично, если не удалось удалить
            }

            [System.Windows.Forms.MessageBox]::Show("Тестовый файл успешно загружен и удалён!`nПодключение к Яндекс.Диску работает корректно.", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Не удалось загрузить тестовый файл:`n$($_.Exception.Message)`n`nПроверьте токен и права доступа.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            # Удаляем локальный тестовый файл
            Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при тестировании Яндекс.Диска:`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Test-TelegramConnection {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$ChatId
    )

    try {
        $telegramModulePath = Join-Path $ProjectRoot 'modules\Notifications.Telegram.psm1'
        if (-not (Test-Path $telegramModulePath)) {
            [System.Windows.Forms.MessageBox]::Show("Модуль Notifications.Telegram.psm1 не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        Import-Module -Force -DisableNameChecking $telegramModulePath -ErrorAction Stop

        $testMessage = "🔧 Тест уведомлений от системы резервного копирования 1С`n" +
                      "📅 Дата и время: {0}`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') +
                      "✅ Статус: Настройка успешно завершена"

        Send-TelegramMessage -Token $Token -ChatId $ChatId -Text $testMessage
        [System.Windows.Forms.MessageBox]::Show("Тестовое сообщение успешно отправлено!", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Не удалось отправить тестовое сообщение:`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# === ЗАГРУЗКА И ЗАПУСК ===
# Загрузка существующих баз в мастер
if (Test-Path $BasesDir) {
    $bases = Get-ChildItem -Path $BasesDir -Filter "*.json" -File -ErrorAction SilentlyContinue
    foreach ($baseFile in $bases) {
        try {
            $config = Get-Content $baseFile.FullName -Raw | ConvertFrom-Json
            $baseObj = [ordered]@{
                Tag = $baseFile.BaseName
                BackupType = if ($config.BackupType) { $config.BackupType } else { "1CD" }
                SourcePath = if ($config.SourcePath) { $config.SourcePath } else { "" }
                DestinationPath = if ($config.DestinationPath) { $config.DestinationPath } else { "" }
                Keep = if ($config.Keep) { [int]$config.Keep } else { 3 }
                CloudType = if ($config.CloudType) { $config.CloudType } else { "" }
                CloudKeep = if ($config.CloudKeep) { [int]$config.CloudKeep } else { 0 }
            }
            $Script:AllBases += $baseObj
        } catch {
            # Игнорируем поврежденные файлы
        }
    }
}

Load-ExistingSettings | Out-Null
Update-WizardProgress
Load-EditorBasesList

# ==========================================
# ВКЛАДКА РАСПИСАНИЕ
# ==========================================

# === ПАНЕЛЬ РАСПИСАНИЯ ===
$SchedulerMainPanel = New-Object System.Windows.Forms.Panel
$SchedulerMainPanel.Location = New-Object System.Drawing.Point(12, 12)
$SchedulerMainPanel.Size = New-Object System.Drawing.Size(1130, 800)
$SchedulerMainPanel.AutoScroll = $true

# === ЗАГОЛОВОК ===
$SchedulerTitleLabel = New-Object System.Windows.Forms.Label
$SchedulerTitleLabel.Text = "Управление расписанием резервного копирования"
$SchedulerTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$SchedulerTitleLabel.Location = New-Object System.Drawing.Point(20, 20)
$SchedulerTitleLabel.Size = New-Object System.Drawing.Size(600, 30)
$SchedulerTitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

$SchedulerInfoLabel = New-Object System.Windows.Forms.Label
$SchedulerInfoLabel.Text = "Настройте автоматическое выполнение резервного копирования по расписанию через Планировщик заданий Windows."
$SchedulerInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$SchedulerInfoLabel.Location = New-Object System.Drawing.Point(20, 55)
$SchedulerInfoLabel.Size = New-Object System.Drawing.Size(900, 20)
$SchedulerInfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen

# === ГРУППА СОЗДАНИЯ РАСПИСАНИЯ ===
$SchedulerGroupBoxCreate = New-Object System.Windows.Forms.GroupBox
$SchedulerGroupBoxCreate.Text = "Создать новое расписание"
$SchedulerGroupBoxCreate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$SchedulerGroupBoxCreate.Location = New-Object System.Drawing.Point(20, 90)
$SchedulerGroupBoxCreate.Size = New-Object System.Drawing.Size(550, 300)

# Имя задачи
$SchedulerLabelTaskName = New-Object System.Windows.Forms.Label
$SchedulerLabelTaskName.Text = "Название задачи:"
$SchedulerLabelTaskName.Location = New-Object System.Drawing.Point(15, 30)
$SchedulerLabelTaskName.Size = New-Object System.Drawing.Size(120, 20)

$SchedulerTextTaskName = New-Object System.Windows.Forms.TextBox
$SchedulerTextTaskName.Location = New-Object System.Drawing.Point(15, 55)
$SchedulerTextTaskName.Size = New-Object System.Drawing.Size(300, 25)
$SchedulerTextTaskName.Text = "Backup1C_Auto_$(Get-Date -Format 'yyyyMMdd')"

# База данных
$SchedulerLabelDatabase = New-Object System.Windows.Forms.Label
$SchedulerLabelDatabase.Text = "База данных:"
$SchedulerLabelDatabase.Location = New-Object System.Drawing.Point(330, 30)
$SchedulerLabelDatabase.Size = New-Object System.Drawing.Size(100, 20)

$SchedulerComboDatabase = New-Object System.Windows.Forms.ComboBox
$SchedulerComboDatabase.Location = New-Object System.Drawing.Point(330, 55)
$SchedulerComboDatabase.Size = New-Object System.Drawing.Size(200, 25)
$SchedulerComboDatabase.DropDownStyle = "DropDownList"
[void]$SchedulerComboDatabase.Items.Add("Все базы")

# Частота
$SchedulerLabelFrequency = New-Object System.Windows.Forms.Label
$SchedulerLabelFrequency.Text = "Частота выполнения:"
$SchedulerLabelFrequency.Location = New-Object System.Drawing.Point(15, 90)
$SchedulerLabelFrequency.Size = New-Object System.Drawing.Size(130, 20)

$SchedulerComboFrequency = New-Object System.Windows.Forms.ComboBox
$SchedulerComboFrequency.Items.AddRange(@("Ежечасно", "Каждые N часов", "Ежедневно", "Еженедельно", "Ежемесячно"))
$SchedulerComboFrequency.DropDownStyle = "DropDownList"
$SchedulerComboFrequency.Location = New-Object System.Drawing.Point(15, 115)
$SchedulerComboFrequency.Size = New-Object System.Drawing.Size(150, 25)
$SchedulerComboFrequency.SelectedIndex = 2  # Default to Daily

# Интервал часов (для "Каждые N часов")
$SchedulerLabelInterval = New-Object System.Windows.Forms.Label
$SchedulerLabelInterval.Text = "Интервал (часы):"
$SchedulerLabelInterval.Location = New-Object System.Drawing.Point(180, 90)
$SchedulerLabelInterval.Size = New-Object System.Drawing.Size(100, 20)
$SchedulerLabelInterval.Visible = $false

$SchedulerNumericInterval = New-Object System.Windows.Forms.NumericUpDown
$SchedulerNumericInterval.Location = New-Object System.Drawing.Point(180, 115)
$SchedulerNumericInterval.Size = New-Object System.Drawing.Size(60, 25)
$SchedulerNumericInterval.Minimum = 1
$SchedulerNumericInterval.Maximum = 24
$SchedulerNumericInterval.Value = 2
$SchedulerNumericInterval.Visible = $false

# Время
$SchedulerLabelTime = New-Object System.Windows.Forms.Label
$SchedulerLabelTime.Text = "Время запуска:"
$SchedulerLabelTime.Location = New-Object System.Drawing.Point(300, 90)
$SchedulerLabelTime.Size = New-Object System.Drawing.Size(100, 20)

$SchedulerDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
$SchedulerDateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$SchedulerDateTimePicker.ShowUpDown = $true
$SchedulerDateTimePicker.Location = New-Object System.Drawing.Point(300, 115)
$SchedulerDateTimePicker.Size = New-Object System.Drawing.Size(100, 25)
$SchedulerDateTimePicker.Value = (Get-Date).Date.AddHours(2)

# Дни недели (для еженедельного расписания)
$SchedulerLabelDays = New-Object System.Windows.Forms.Label
$SchedulerLabelDays.Text = "Дни недели:"
$SchedulerLabelDays.Location = New-Object System.Drawing.Point(15, 150)
$SchedulerLabelDays.Size = New-Object System.Drawing.Size(100, 20)
$SchedulerLabelDays.Visible = $false

$SchedulerCheckListDays = New-Object System.Windows.Forms.CheckedListBox
$SchedulerCheckListDays.Items.AddRange(@("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"))
$SchedulerCheckListDays.Location = New-Object System.Drawing.Point(15, 175)
$SchedulerCheckListDays.Size = New-Object System.Drawing.Size(300, 90)
$SchedulerCheckListDays.Visible = $false
$SchedulerCheckListDays.SetItemChecked(0, $true) # Понедельник по умолчанию
$SchedulerCheckListDays.SetItemChecked(2, $true) # Среда
$SchedulerCheckListDays.SetItemChecked(4, $true) # Пятница

# День месяца (для ежемесячного расписания)
$SchedulerLabelDayOfMonth = New-Object System.Windows.Forms.Label
$SchedulerLabelDayOfMonth.Text = "День месяца:"
$SchedulerLabelDayOfMonth.Location = New-Object System.Drawing.Point(300, 90)
$SchedulerLabelDayOfMonth.Size = New-Object System.Drawing.Size(100, 20)
$SchedulerLabelDayOfMonth.Visible = $false

$SchedulerNumericDayOfMonth = New-Object System.Windows.Forms.NumericUpDown
$SchedulerNumericDayOfMonth.Location = New-Object System.Drawing.Point(300, 115)
$SchedulerNumericDayOfMonth.Size = New-Object System.Drawing.Size(60, 25)
$SchedulerNumericDayOfMonth.Minimum = 1
$SchedulerNumericDayOfMonth.Maximum = 31
$SchedulerNumericDayOfMonth.Value = 1
$SchedulerNumericDayOfMonth.Visible = $false

# Кнопки
$SchedulerButtonCreate = New-Object System.Windows.Forms.Button
$SchedulerButtonCreate.Text = "Создать расписание"
$SchedulerButtonCreate.Location = New-Object System.Drawing.Point(15, 270)
$SchedulerButtonCreate.Size = New-Object System.Drawing.Size(150, 25)
$SchedulerButtonCreate.BackColor = [System.Drawing.Color]::LightGreen

$SchedulerButtonTestRun = New-Object System.Windows.Forms.Button
$SchedulerButtonTestRun.Text = "Тестовый запуск"
$SchedulerButtonTestRun.Location = New-Object System.Drawing.Point(175, 270)
$SchedulerButtonTestRun.Size = New-Object System.Drawing.Size(120, 25)

$SchedulerGroupBoxCreate.Controls.AddRange(@(
    $SchedulerLabelTaskName, $SchedulerTextTaskName,
    $SchedulerLabelDatabase, $SchedulerComboDatabase,
    $SchedulerLabelFrequency, $SchedulerComboFrequency,
    $SchedulerLabelInterval, $SchedulerNumericInterval,
    $SchedulerLabelTime, $SchedulerDateTimePicker,
    $SchedulerLabelDays, $SchedulerCheckListDays,
    $SchedulerLabelDayOfMonth, $SchedulerNumericDayOfMonth,
    $SchedulerButtonCreate, $SchedulerButtonTestRun
))

# === ГРУППА АКТИВНЫХ ЗАДАЧ ===
$SchedulerGroupBoxTasks = New-Object System.Windows.Forms.GroupBox
$SchedulerGroupBoxTasks.Text = "Активные задачи"
$SchedulerGroupBoxTasks.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$SchedulerGroupBoxTasks.Location = New-Object System.Drawing.Point(590, 90)
$SchedulerGroupBoxTasks.Size = New-Object System.Drawing.Size(520, 300)

$SchedulerListTasks = New-Object System.Windows.Forms.ListView
$SchedulerListTasks.Location = New-Object System.Drawing.Point(15, 25)
$SchedulerListTasks.Size = New-Object System.Drawing.Size(490, 200)
$SchedulerListTasks.View = "Details"
$SchedulerListTasks.FullRowSelect = $true
$SchedulerListTasks.GridLines = $true
[void]$SchedulerListTasks.Columns.Add("Название", 150)
[void]$SchedulerListTasks.Columns.Add("Состояние", 80)
[void]$SchedulerListTasks.Columns.Add("Последний запуск", 120)
[void]$SchedulerListTasks.Columns.Add("Следующий запуск", 120)

$SchedulerButtonRefresh = New-Object System.Windows.Forms.Button
$SchedulerButtonRefresh.Text = "Обновить"
$SchedulerButtonRefresh.Location = New-Object System.Drawing.Point(15, 235)
$SchedulerButtonRefresh.Size = New-Object System.Drawing.Size(80, 25)

$SchedulerButtonRun = New-Object System.Windows.Forms.Button
$SchedulerButtonRun.Text = "Запустить"
$SchedulerButtonRun.Location = New-Object System.Drawing.Point(105, 235)
$SchedulerButtonRun.Size = New-Object System.Drawing.Size(80, 25)
$SchedulerButtonRun.BackColor = [System.Drawing.Color]::LightBlue

$SchedulerButtonDelete = New-Object System.Windows.Forms.Button
$SchedulerButtonDelete.Text = "Удалить"
$SchedulerButtonDelete.Location = New-Object System.Drawing.Point(195, 235)
$SchedulerButtonDelete.Size = New-Object System.Drawing.Size(80, 25)
$SchedulerButtonDelete.BackColor = [System.Drawing.Color]::LightCoral

$SchedulerGroupBoxTasks.Controls.AddRange(@(
    $SchedulerListTasks,
    $SchedulerButtonRefresh, $SchedulerButtonRun, $SchedulerButtonDelete
))

$SchedulerMainPanel.Controls.AddRange(@(
    $SchedulerTitleLabel, $SchedulerInfoLabel,
    $SchedulerGroupBoxCreate, $SchedulerGroupBoxTasks
))

$TabScheduler.Controls.Add($SchedulerMainPanel)

# === ОБРАБОТЧИКИ СОБЫТИЙ РАСПИСАНИЯ ===

# Обработчик изменения частоты
$SchedulerComboFrequency.Add_SelectedIndexChanged({
    switch ($SchedulerComboFrequency.SelectedIndex) {
        0 { # Ежечасно
            $SchedulerLabelInterval.Visible = $false
            $SchedulerNumericInterval.Visible = $false
            $SchedulerLabelTime.Visible = $false
            $SchedulerDateTimePicker.Visible = $false
            $SchedulerLabelDays.Visible = $false
            $SchedulerCheckListDays.Visible = $false
            $SchedulerLabelDayOfMonth.Visible = $false
            $SchedulerNumericDayOfMonth.Visible = $false
        }
        1 { # Каждые N часов
            $SchedulerLabelInterval.Visible = $true
            $SchedulerNumericInterval.Visible = $true
            $SchedulerLabelTime.Visible = $true
            $SchedulerDateTimePicker.Visible = $true
            $SchedulerLabelDays.Visible = $false
            $SchedulerCheckListDays.Visible = $false
            $SchedulerLabelDayOfMonth.Visible = $false
            $SchedulerNumericDayOfMonth.Visible = $false
        }
        2 { # Ежедневно
            $SchedulerLabelInterval.Visible = $false
            $SchedulerNumericInterval.Visible = $false
            $SchedulerLabelTime.Visible = $true
            $SchedulerDateTimePicker.Visible = $true
            $SchedulerLabelDays.Visible = $false
            $SchedulerCheckListDays.Visible = $false
            $SchedulerLabelDayOfMonth.Visible = $false
            $SchedulerNumericDayOfMonth.Visible = $false
        }
        3 { # Еженедельно
            $SchedulerLabelInterval.Visible = $false
            $SchedulerNumericInterval.Visible = $false
            $SchedulerLabelTime.Visible = $true
            $SchedulerDateTimePicker.Visible = $true
            $SchedulerLabelDays.Visible = $true
            $SchedulerCheckListDays.Visible = $true
            $SchedulerLabelDayOfMonth.Visible = $false
            $SchedulerNumericDayOfMonth.Visible = $false
        }
        4 { # Ежемесячно
            $SchedulerLabelInterval.Visible = $false
            $SchedulerNumericInterval.Visible = $false
            $SchedulerLabelTime.Visible = $true
            $SchedulerDateTimePicker.Visible = $true
            $SchedulerLabelDays.Visible = $false
            $SchedulerCheckListDays.Visible = $false
            $SchedulerLabelDayOfMonth.Visible = $true
            $SchedulerNumericDayOfMonth.Visible = $true
        }
    }
})

# Функция обновления списка задач
function Refresh-SchedulerTasksList {
    try {
        $SchedulerListTasks.Items.Clear()
        $tasksResult = Get-AllBackupScheduledTasks

        if ($tasksResult.Success) {
            foreach ($task in $tasksResult.Tasks) {
                try {
                    $item = New-Object System.Windows.Forms.ListViewItem($task.Name)
                    [void]$item.SubItems.Add($task.State)
                    $lastRun = if ($task.LastRunTime) { $task.LastRunTime.ToString("yyyy-MM-dd HH:mm") } else { "Никогда" }
                    [void]$item.SubItems.Add($lastRun)
                    $nextRun = if ($task.NextRunTime) { $task.NextRunTime.ToString("yyyy-MM-dd HH:mm") } else { "Не запланирован" }
                    [void]$item.SubItems.Add($nextRun)
                    $item.Tag = $task
                    [void]$SchedulerListTasks.Items.Add($item)
                }
                catch {
                    # Skip this task if there's an error
                    Write-Host "Ошибка добавления задачи $($task.Name): $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Не удалось получить список задач: $($tasksResult.Message)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Критическая ошибка в Refresh-SchedulerTasksList: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Обработчик создания расписания
$SchedulerButtonCreate.Add_Click({
    try {
        $taskName = $SchedulerTextTaskName.Text.Trim()
        if ([string]::IsNullOrEmpty($taskName)) {
            [System.Windows.Forms.MessageBox]::Show("Введите название задачи!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $frequency = switch ($SchedulerComboFrequency.SelectedIndex) {
            0 { "Hourly" }        # Ежечасно
            1 { "CustomHours" }   # Каждые N часов
            2 { "Daily" }         # Ежедневно
            3 { "Weekly" }        # Еженедельно
            4 { "Monthly" }       # Ежемесячно
            default { "Daily" }
        }

        $time = $SchedulerDateTimePicker.Value.ToString("HH:mm")
        $scriptPath = Join-Path $ProjectRoot "Run-Backup.ps1"

        # Check if specific database is selected
        $selectedDatabase = $SchedulerComboDatabase.SelectedItem
        if ($selectedDatabase -and $selectedDatabase -ne "Все базы") {
            $scriptArguments = "-ExecutionPolicy Bypass -File `"$scriptPath`" -BaseName `"$selectedDatabase`""
            $description = "Автоматическое резервное копирование базы '$selectedDatabase' - создано через Backup1C Manager"
        } else {
            $scriptArguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
            $description = "Автоматическое резервное копирование всех баз 1С - создано через Backup1C Manager"
        }

        $params = @{
            TaskName = $taskName
            ScriptPath = $scriptPath
            ScriptArguments = $scriptArguments
            Frequency = $frequency
            Time = $time
            Description = $description
        }

        if ($frequency -eq "CustomHours") {
            $params.Interval = [int]$SchedulerNumericInterval.Value
        } elseif ($frequency -eq "Weekly") {
            $selectedDays = @()
            $dayMapping = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
            for ($i = 0; $i -lt $SchedulerCheckListDays.Items.Count; $i++) {
                if ($SchedulerCheckListDays.GetItemChecked($i)) {
                    $selectedDays += $dayMapping[$i]
                }
            }

            if ($selectedDays.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Выберите хотя бы один день недели!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $params.Days = $selectedDays -join ","
        } elseif ($frequency -eq "Monthly") {
            $params.DayOfMonth = [int]$SchedulerNumericDayOfMonth.Value
        }

        $result = New-BackupScheduledTask @params

        if ($result.Success) {
            [System.Windows.Forms.MessageBox]::Show($result.Message, "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Refresh-SchedulerTasksList
        } else {
            [System.Windows.Forms.MessageBox]::Show($result.Message, "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка создания расписания: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Обработчик тестового запуска
$SchedulerButtonTestRun.Add_Click({
    try {
        $scriptPath = Join-Path $ProjectRoot "Run-Backup.ps1"
        if (Test-Path $scriptPath) {
            Start-Process -FilePath "PowerShell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Normal
            [System.Windows.Forms.MessageBox]::Show("Тестовое резервное копирование запущено в отдельном окне.", "Запуск", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Файл Run-Backup.ps1 не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка запуска: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Обработчик обновления списка задач
$SchedulerButtonRefresh.Add_Click({
    try {
        Refresh-SchedulerTasksList
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка обновления списка задач: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Обработчик запуска выбранной задачи
$SchedulerButtonRun.Add_Click({
    if ($SchedulerListTasks.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Выберите задачу для запуска!", "Предупреждение", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedTask = $SchedulerListTasks.SelectedItems[0]
    $taskName = $selectedTask.Text

    $result = Start-BackupScheduledTask -TaskName $taskName
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Refresh-SchedulerTasksList
    } else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Обработчик удаления задачи
$SchedulerButtonDelete.Add_Click({
    if ($SchedulerListTasks.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Выберите задачу для удаления!", "Предупреждение", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedTask = $SchedulerListTasks.SelectedItems[0]
    $taskName = $selectedTask.Text

    $confirmResult = [System.Windows.Forms.MessageBox]::Show("Удалить задачу '$taskName'?", "Подтверждение", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $result = Remove-BackupScheduledTask -TaskName $taskName
        if ($result.Success) {
            [System.Windows.Forms.MessageBox]::Show($result.Message, "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Refresh-SchedulerTasksList
        } else {
            [System.Windows.Forms.MessageBox]::Show($result.Message, "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Загрузка списка задач при запуске
Refresh-SchedulerTasksList

# ==========================================
# ВКЛАДКА МОНИТОРИНГ
# ==========================================

# === ПАНЕЛЬ МОНИТОРИНГА ===
$MonitorMainPanel = New-Object System.Windows.Forms.Panel
$MonitorMainPanel.Location = New-Object System.Drawing.Point(12, 12)
$MonitorMainPanel.Size = New-Object System.Drawing.Size(1130, 800)
$MonitorMainPanel.AutoScroll = $true

# === ЗАГОЛОВОК ===
$MonitorTitleLabel = New-Object System.Windows.Forms.Label
$MonitorTitleLabel.Text = "Мониторинг системы резервного копирования"
$MonitorTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$MonitorTitleLabel.Location = New-Object System.Drawing.Point(20, 20)
$MonitorTitleLabel.Size = New-Object System.Drawing.Size(600, 30)
$MonitorTitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

# === ГРУППА БЫСТРЫХ ДЕЙСТВИЙ ===
$MonitorGroupBoxActions = New-Object System.Windows.Forms.GroupBox
$MonitorGroupBoxActions.Text = "Быстрые действия"
$MonitorGroupBoxActions.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$MonitorGroupBoxActions.Location = New-Object System.Drawing.Point(20, 60)
$MonitorGroupBoxActions.Size = New-Object System.Drawing.Size(1090, 80)

$MonitorButtonRunBackup = New-Object System.Windows.Forms.Button
$MonitorButtonRunBackup.Text = "Запустить резервное копирование"
$MonitorButtonRunBackup.Location = New-Object System.Drawing.Point(15, 25)
$MonitorButtonRunBackup.Size = New-Object System.Drawing.Size(200, 35)
$MonitorButtonRunBackup.BackColor = [System.Drawing.Color]::LightGreen
$MonitorButtonRunBackup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$MonitorButtonViewLogs = New-Object System.Windows.Forms.Button
$MonitorButtonViewLogs.Text = "Открыть папку с логами"
$MonitorButtonViewLogs.Location = New-Object System.Drawing.Point(230, 25)
$MonitorButtonViewLogs.Size = New-Object System.Drawing.Size(180, 35)
$MonitorButtonViewLogs.BackColor = [System.Drawing.Color]::LightBlue

$MonitorButtonOpenTaskScheduler = New-Object System.Windows.Forms.Button
$MonitorButtonOpenTaskScheduler.Text = "Планировщик заданий"
$MonitorButtonOpenTaskScheduler.Location = New-Object System.Drawing.Point(425, 25)
$MonitorButtonOpenTaskScheduler.Size = New-Object System.Drawing.Size(160, 35)

$MonitorButtonRefreshAll = New-Object System.Windows.Forms.Button
$MonitorButtonRefreshAll.Text = "Обновить все"
$MonitorButtonRefreshAll.Location = New-Object System.Drawing.Point(600, 25)
$MonitorButtonRefreshAll.Size = New-Object System.Drawing.Size(120, 35)
$MonitorButtonRefreshAll.BackColor = [System.Drawing.Color]::LightYellow

$MonitorGroupBoxActions.Controls.AddRange(@(
    $MonitorButtonRunBackup, $MonitorButtonViewLogs,
    $MonitorButtonOpenTaskScheduler, $MonitorButtonRefreshAll
))

# === ГРУППА СТАТУСА ===
$MonitorGroupBoxStatus = New-Object System.Windows.Forms.GroupBox
$MonitorGroupBoxStatus.Text = "Текущий статус"
$MonitorGroupBoxStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$MonitorGroupBoxStatus.Location = New-Object System.Drawing.Point(20, 160)
$MonitorGroupBoxStatus.Size = New-Object System.Drawing.Size(530, 200)

$MonitorLabelLastBackup = New-Object System.Windows.Forms.Label
$MonitorLabelLastBackup.Text = "Последнее резервное копирование:"
$MonitorLabelLastBackup.Location = New-Object System.Drawing.Point(15, 25)
$MonitorLabelLastBackup.Size = New-Object System.Drawing.Size(200, 20)
$MonitorLabelLastBackup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$MonitorValueLastBackup = New-Object System.Windows.Forms.Label
$MonitorValueLastBackup.Text = "Проверка..."
$MonitorValueLastBackup.Location = New-Object System.Drawing.Point(220, 25)
$MonitorValueLastBackup.Size = New-Object System.Drawing.Size(300, 20)
$MonitorValueLastBackup.ForeColor = [System.Drawing.Color]::DarkGreen

$MonitorLabelActiveTasks = New-Object System.Windows.Forms.Label
$MonitorLabelActiveTasks.Text = "Активных задач в планировщике:"
$MonitorLabelActiveTasks.Location = New-Object System.Drawing.Point(15, 55)
$MonitorLabelActiveTasks.Size = New-Object System.Drawing.Size(200, 20)
$MonitorLabelActiveTasks.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$MonitorValueActiveTasks = New-Object System.Windows.Forms.Label
$MonitorValueActiveTasks.Text = "Проверка..."
$MonitorValueActiveTasks.Location = New-Object System.Drawing.Point(220, 55)
$MonitorValueActiveTasks.Size = New-Object System.Drawing.Size(300, 20)
$MonitorValueActiveTasks.ForeColor = [System.Drawing.Color]::DarkBlue

$MonitorLabelDiskSpace = New-Object System.Windows.Forms.Label
$MonitorLabelDiskSpace.Text = "Свободное место на дисках:"
$MonitorLabelDiskSpace.Location = New-Object System.Drawing.Point(15, 85)
$MonitorLabelDiskSpace.Size = New-Object System.Drawing.Size(200, 20)
$MonitorLabelDiskSpace.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$MonitorValueDiskSpace = New-Object System.Windows.Forms.Label
$MonitorValueDiskSpace.Text = "Проверка..."
$MonitorValueDiskSpace.Location = New-Object System.Drawing.Point(15, 110)
$MonitorValueDiskSpace.Size = New-Object System.Drawing.Size(500, 60)
$MonitorValueDiskSpace.ForeColor = [System.Drawing.Color]::DarkGray

$MonitorGroupBoxStatus.Controls.AddRange(@(
    $MonitorLabelLastBackup, $MonitorValueLastBackup,
    $MonitorLabelActiveTasks, $MonitorValueActiveTasks,
    $MonitorLabelDiskSpace, $MonitorValueDiskSpace
))

# === ГРУППА НЕДАВНИХ ЛОГОВ ===
$MonitorGroupBoxLogs = New-Object System.Windows.Forms.GroupBox
$MonitorGroupBoxLogs.Text = "Недавние логи"
$MonitorGroupBoxLogs.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$MonitorGroupBoxLogs.Location = New-Object System.Drawing.Point(570, 160)
$MonitorGroupBoxLogs.Size = New-Object System.Drawing.Size(540, 400)

$MonitorListLogs = New-Object System.Windows.Forms.ListBox
$MonitorListLogs.Location = New-Object System.Drawing.Point(15, 25)
$MonitorListLogs.Size = New-Object System.Drawing.Size(510, 340)
$MonitorListLogs.Font = New-Object System.Drawing.Font("Consolas", 8)
$MonitorListLogs.ScrollAlwaysVisible = $true

$MonitorButtonClearLogs = New-Object System.Windows.Forms.Button
$MonitorButtonClearLogs.Text = "Очистить отображение"
$MonitorButtonClearLogs.Location = New-Object System.Drawing.Point(15, 370)
$MonitorButtonClearLogs.Size = New-Object System.Drawing.Size(150, 25)

$MonitorGroupBoxLogs.Controls.AddRange(@($MonitorListLogs, $MonitorButtonClearLogs))

$MonitorMainPanel.Controls.AddRange(@(
    $MonitorTitleLabel, $MonitorGroupBoxActions,
    $MonitorGroupBoxStatus, $MonitorGroupBoxLogs
))

$TabMonitor.Controls.Add($MonitorMainPanel)

# === ОБРАБОТЧИКИ СОБЫТИЙ МОНИТОРИНГА ===

# Функция обновления статуса
function Update-MonitorStatus {
    # Проверка последнего резервного копирования
    $logsDir = Join-Path $ProjectRoot "logs"
    if (Test-Path $logsDir) {
        $latestLog = Get-ChildItem -Path $logsDir -Filter "backup_*.log" -File | Sort-Object CreationTime -Descending | Select-Object -First 1
        if ($latestLog) {
            $MonitorValueLastBackup.Text = "$($latestLog.CreationTime.ToString('yyyy-MM-dd HH:mm'))"
            $MonitorValueLastBackup.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            $MonitorValueLastBackup.Text = "Логи не найдены"
            $MonitorValueLastBackup.ForeColor = [System.Drawing.Color]::DarkRed
        }
    } else {
        $MonitorValueLastBackup.Text = "Папка логов не найдена"
        $MonitorValueLastBackup.ForeColor = [System.Drawing.Color]::DarkRed
    }

    # Проверка активных задач
    $tasksResult = Get-AllBackupScheduledTasks
    if ($tasksResult.Success) {
        $activeTasks = $tasksResult.Tasks | Where-Object { $_.State -eq "Ready" }
        $MonitorValueActiveTasks.Text = "$($activeTasks.Count) из $($tasksResult.Tasks.Count) готовы к выполнению"
        $MonitorValueActiveTasks.ForeColor = if ($activeTasks.Count -gt 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed }
    } else {
        $MonitorValueActiveTasks.Text = "Ошибка получения задач"
        $MonitorValueActiveTasks.ForeColor = [System.Drawing.Color]::DarkRed
    }

    # Проверка свободного места на дисках
    try {
        $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $diskInfo = @()
        foreach ($drive in $drives) {
            $freeGB = [math]::Round($drive.FreeSpace / 1GB, 1)
            $totalGB = [math]::Round($drive.Size / 1GB, 1)
            $freePercent = [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1)
            $diskInfo += "Диск $($drive.DeviceID) $freeGB/$totalGB ГБ ($freePercent%)"
        }
        $MonitorValueDiskSpace.Text = $diskInfo -join "`n"
        $MonitorValueDiskSpace.ForeColor = [System.Drawing.Color]::DarkGray
    } catch {
        $MonitorValueDiskSpace.Text = "Ошибка получения информации о дисках"
        $MonitorValueDiskSpace.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

# Обработчики кнопок мониторинга
$MonitorButtonRunBackup.Add_Click({
    try {
        $scriptPath = Join-Path $ProjectRoot "Run-Backup.ps1"
        if (Test-Path $scriptPath) {
            Start-Process -FilePath "PowerShell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Normal
            [void]$MonitorListLogs.Items.Add("$(Get-Date -Format 'HH:mm:ss') - Запущено ручное резервное копирование")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Файл Run-Backup.ps1 не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка запуска: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$MonitorButtonViewLogs.Add_Click({
    $logsDir = Join-Path $ProjectRoot "logs"
    if (Test-Path $logsDir) {
        Start-Process -FilePath "explorer.exe" -ArgumentList $logsDir
    } else {
        [System.Windows.Forms.MessageBox]::Show("Папка логов не найдена: $logsDir", "Предупреждение", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$MonitorButtonOpenTaskScheduler.Add_Click({
    try {
        Start-Process -FilePath "taskschd.msc"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка открытия планировщика заданий: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$MonitorButtonRefreshAll.Add_Click({
    Update-MonitorStatus
    Refresh-SchedulerTasksList
    [void]$MonitorListLogs.Items.Add("$(Get-Date -Format 'HH:mm:ss') - Обновлена информация о статусе системы")
})

$MonitorButtonClearLogs.Add_Click({
    $MonitorListLogs.Items.Clear()
})

# Инициализация мониторинга
Update-MonitorStatus
[void]$MonitorListLogs.Items.Add("$(Get-Date -Format 'HH:mm:ss') - Система мониторинга запущена")

try {
    [void]$MainForm.ShowDialog()
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Критическая ошибка: $_",
        "Ошибка",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

$MainForm.Dispose()