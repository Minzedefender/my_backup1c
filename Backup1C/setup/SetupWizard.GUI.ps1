#requires -Version 5.1
# Графический мастер первичной настройки системы бэкапа 1С

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Import-Module -Force -DisableNameChecking (Join-Path $PSScriptRoot '..\modules\Common.Crypto.psm1')

# Глобальные переменные
$Script:AllBases = @()
$Script:AllSecrets = @{}
$Script:ConfigRoot = Join-Path $PSScriptRoot '..\config'
$Script:BasesDir = Join-Path $Script:ConfigRoot 'bases'
$Script:SettingsFile = Join-Path $Script:ConfigRoot 'settings.json'
$Script:KeyPath = Join-Path $Script:ConfigRoot 'key.bin'
$Script:SecretsFile = Join-Path $Script:ConfigRoot 'secrets.json.enc'
$Script:CurrentStep = 1
$Script:MaxSteps = 3

# Создание директорий
if (-not (Test-Path $Script:ConfigRoot)) { New-Item -ItemType Directory -Path $Script:ConfigRoot -Force | Out-Null }
if (-not (Test-Path $Script:BasesDir)) { New-Item -ItemType Directory -Path $Script:BasesDir -Force | Out-Null }

# Загрузка существующих настроек
function Load-ExistingSettings {
    # Загрузка секретов
    if ((Test-Path $Script:SecretsFile) -and (Test-Path $Script:KeyPath)) {
        try {
            $raw = Decrypt-Secrets -InFile $Script:SecretsFile -KeyPath $Script:KeyPath
            $Script:AllSecrets = ConvertTo-Hashtable $raw
            if (-not ($Script:AllSecrets -is [hashtable])) { $Script:AllSecrets = @{} }
        } catch { $Script:AllSecrets = @{} }
    }

    # Загрузка настроек
    $settingsData = @{}
    if (Test-Path $Script:SettingsFile) {
        try {
            $settingsData = ConvertTo-Hashtable (Get-Content $Script:SettingsFile -Raw | ConvertFrom-Json)
            if (-not ($settingsData -is [hashtable])) { $settingsData = @{} }
        } catch { $settingsData = @{} }
    }

    return $settingsData
}

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

# === ГЛАВНАЯ ФОРМА ===
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Мастер настройки системы резервного копирования 1С"
$MainForm.Size = New-Object System.Drawing.Size(1000, 700)
$MainForm.StartPosition = "CenterScreen"
$MainForm.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$MainForm.FormBorderStyle = "FixedDialog"
$MainForm.MaximizeBox = $false

# === ПРОГРЕСС-БАР ===
$ProgressPanel = New-Object System.Windows.Forms.Panel
$ProgressPanel.Location = New-Object System.Drawing.Point(12, 12)
$ProgressPanel.Size = New-Object System.Drawing.Size(960, 60)
$ProgressPanel.BackColor = [System.Drawing.Color]::WhiteSmoke

$ProgressLabel = New-Object System.Windows.Forms.Label
$ProgressLabel.Location = New-Object System.Drawing.Point(10, 10)
$ProgressLabel.Size = New-Object System.Drawing.Size(940, 20)
$ProgressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$ProgressLabel.Text = "Шаг 1 из 3: Настройка баз данных"

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(10, 35)
$ProgressBar.Size = New-Object System.Drawing.Size(940, 20)
$ProgressBar.Maximum = $Script:MaxSteps
$ProgressBar.Value = $Script:CurrentStep

$ProgressPanel.Controls.AddRange(@($ProgressLabel, $ProgressBar))

# === ПАНЕЛЬ КОНТЕНТА ===
$ContentPanel = New-Object System.Windows.Forms.Panel
$ContentPanel.Location = New-Object System.Drawing.Point(12, 85)
$ContentPanel.Size = New-Object System.Drawing.Size(960, 520)
$ContentPanel.BackColor = [System.Drawing.Color]::White
$ContentPanel.BorderStyle = "FixedSingle"

# === ПАНЕЛЬ КНОПОК ===
$ButtonPanel = New-Object System.Windows.Forms.Panel
$ButtonPanel.Location = New-Object System.Drawing.Point(12, 615)
$ButtonPanel.Size = New-Object System.Drawing.Size(960, 50)

$ButtonBack = New-Object System.Windows.Forms.Button
$ButtonBack.Text = "< Назад"
$ButtonBack.Location = New-Object System.Drawing.Point(10, 10)
$ButtonBack.Size = New-Object System.Drawing.Size(100, 35)
$ButtonBack.Enabled = $false

$ButtonNext = New-Object System.Windows.Forms.Button
$ButtonNext.Text = "Далее >"
$ButtonNext.Location = New-Object System.Drawing.Point(750, 10)
$ButtonNext.Size = New-Object System.Drawing.Size(100, 35)
$ButtonNext.UseVisualStyleBackColor = $true

$ButtonFinish = New-Object System.Windows.Forms.Button
$ButtonFinish.Text = "Завершить"
$ButtonFinish.Location = New-Object System.Drawing.Point(860, 10)
$ButtonFinish.Size = New-Object System.Drawing.Size(100, 35)
$ButtonFinish.BackColor = [System.Drawing.Color]::LightGreen
$ButtonFinish.UseVisualStyleBackColor = $false
$ButtonFinish.Visible = $false

$ButtonPanel.Controls.AddRange(@($ButtonBack, $ButtonNext, $ButtonFinish))

# Добавление всех панелей на главную форму
$MainForm.Controls.AddRange(@($ProgressPanel, $ContentPanel, $ButtonPanel))

# === ШАГ 1: НАСТРОЙКА БАЗ ДАННЫХ ===
function Show-Step1 {
    $ContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Настройка баз данных для резервного копирования"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(900, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Список баз
    $BasesGroupBox = New-Object System.Windows.Forms.GroupBox
    $BasesGroupBox.Text = "Список баз для резервного копирования"
    $BasesGroupBox.Location = New-Object System.Drawing.Point(20, 60)
    $BasesGroupBox.Size = New-Object System.Drawing.Size(920, 300)

    $Script:BasesListView = New-Object System.Windows.Forms.ListView
    $Script:BasesListView.View = "Details"
    $Script:BasesListView.FullRowSelect = $true
    $Script:BasesListView.GridLines = $true
    $Script:BasesListView.Location = New-Object System.Drawing.Point(10, 25)
    $Script:BasesListView.Size = New-Object System.Drawing.Size(900, 230)

    $Script:BasesListView.Columns.Add("Имя базы", 150) | Out-Null
    $Script:BasesListView.Columns.Add("Тип", 60) | Out-Null
    $Script:BasesListView.Columns.Add("Путь к базе", 300) | Out-Null
    $Script:BasesListView.Columns.Add("Папка бэкапов", 250) | Out-Null
    $Script:BasesListView.Columns.Add("Копий", 60) | Out-Null
    $Script:BasesListView.Columns.Add("Облако", 80) | Out-Null

    $AddBaseButton = New-Object System.Windows.Forms.Button
    $AddBaseButton.Text = "Добавить базу"
    $AddBaseButton.Location = New-Object System.Drawing.Point(10, 265)
    $AddBaseButton.Size = New-Object System.Drawing.Size(120, 30)
    $AddBaseButton.BackColor = [System.Drawing.Color]::LightGreen

    $EditBaseButton = New-Object System.Windows.Forms.Button
    $EditBaseButton.Text = "Редактировать"
    $EditBaseButton.Location = New-Object System.Drawing.Point(140, 265)
    $EditBaseButton.Size = New-Object System.Drawing.Size(120, 30)
    $EditBaseButton.Enabled = $false

    $DeleteBaseButton = New-Object System.Windows.Forms.Button
    $DeleteBaseButton.Text = "Удалить"
    $DeleteBaseButton.Location = New-Object System.Drawing.Point(270, 265)
    $DeleteBaseButton.Size = New-Object System.Drawing.Size(120, 30)
    $DeleteBaseButton.BackColor = [System.Drawing.Color]::LightCoral
    $DeleteBaseButton.Enabled = $false

    $BasesGroupBox.Controls.AddRange(@($Script:BasesListView, $AddBaseButton, $EditBaseButton, $DeleteBaseButton))

    # Информационная панель
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Text = "Добавьте базы данных 1С, для которых нужно настроить резервное копирование.`nВы можете настроить копирование файлов .1CD или выгрузку в формат .dt через конфигуратор."
    $InfoLabel.Location = New-Object System.Drawing.Point(20, 380)
    $InfoLabel.Size = New-Object System.Drawing.Size(900, 50)
    $InfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $InfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen

    $ContentPanel.Controls.AddRange(@($TitleLabel, $BasesGroupBox, $InfoLabel))

    # Обработчики событий
    $Script:BasesListView.Add_SelectedIndexChanged({
        $hasSelection = $Script:BasesListView.SelectedItems.Count -gt 0
        $EditBaseButton.Enabled = $hasSelection
        $DeleteBaseButton.Enabled = $hasSelection
    })

    $AddBaseButton.Add_Click({ Show-BaseEditDialog })
    $EditBaseButton.Add_Click({ Show-BaseEditDialog -EditMode $true })
    $DeleteBaseButton.Add_Click({
        if ($Script:BasesListView.SelectedItems.Count -gt 0) {
            $selectedIndex = $Script:BasesListView.SelectedItems[0].Index
            $baseName = $Script:AllBases[$selectedIndex].Tag
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Удалить базу '$baseName' из настроек?",
                "Подтверждение",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                $Script:AllBases = $Script:AllBases | Where-Object { $_.Tag -ne $baseName }
                Update-BasesListView
            }
        }
    })

    Update-BasesListView
}

function Update-BasesListView {
    $Script:BasesListView.Items.Clear()
    foreach ($base in $Script:AllBases) {
        $item = New-Object System.Windows.Forms.ListViewItem($base.Tag)
        $item.SubItems.Add($base.BackupType) | Out-Null
        $item.SubItems.Add($base.SourcePath) | Out-Null
        $item.SubItems.Add($base.DestinationPath) | Out-Null
        $item.SubItems.Add($base.Keep.ToString()) | Out-Null
        $cloudText = if ($base.CloudType -eq 'Yandex.Disk') { 'Да' } else { 'Нет' }
        $item.SubItems.Add($cloudText) | Out-Null
        $Script:BasesListView.Items.Add($item) | Out-Null
    }
}

function Show-BaseEditDialog {
    param([bool]$EditMode = $false)

    $editBase = $null
    $editIndex = -1

    if ($EditMode -and $Script:BasesListView.SelectedItems.Count -gt 0) {
        $editIndex = $Script:BasesListView.SelectedItems[0].Index
        $editBase = $Script:AllBases[$editIndex]
    }

    $BaseDialog = New-Object System.Windows.Forms.Form
    $BaseDialog.Text = if ($EditMode) { "Редактирование базы" } else { "Добавление новой базы" }
    $BaseDialog.Size = New-Object System.Drawing.Size(600, 820)
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
    $GroupBoxDT.Size = New-Object System.Drawing.Size(550, 250)
    $GroupBoxDT.Enabled = ($ComboType.Text -eq "DT")

    # 1cestart.exe
    $LabelExe = New-Object System.Windows.Forms.Label
    $LabelExe.Text = "Путь к 1cestart.exe (обязательно):"
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

    $CheckBoxAutoServices = New-Object System.Windows.Forms.RadioButton
    $CheckBoxAutoServices.Text = "Автоматически (Apache2.4 или все веб-службы)"
    $CheckBoxAutoServices.Location = New-Object System.Drawing.Point(10, 170)
    $CheckBoxAutoServices.Size = New-Object System.Drawing.Size(350, 20)
    $CheckBoxAutoServices.Checked = $true

    $CheckBoxManualServices = New-Object System.Windows.Forms.RadioButton
    $CheckBoxManualServices.Text = "Выбрать вручную"
    $CheckBoxManualServices.Location = New-Object System.Drawing.Point(10, 195)
    $CheckBoxManualServices.Size = New-Object System.Drawing.Size(150, 20)

    $CheckBoxNoServices = New-Object System.Windows.Forms.RadioButton
    $CheckBoxNoServices.Text = "Не останавливать"
    $CheckBoxNoServices.Location = New-Object System.Drawing.Point(10, 215)
    $CheckBoxNoServices.Size = New-Object System.Drawing.Size(150, 20)

    $GroupBoxDT.Controls.AddRange(@($LabelExe, $TextExe, $ButtonBrowseExe, $LabelDTLogin, $TextDTLogin, $LabelDTPassword, $TextDTPassword, $LabelServices, $CheckBoxAutoServices, $CheckBoxManualServices, $CheckBoxNoServices))

    # Облачное хранилище
    $GroupBoxCloud = New-Object System.Windows.Forms.GroupBox
    $GroupBoxCloud.Text = "Яндекс.Диск"
    $GroupBoxCloud.Location = New-Object System.Drawing.Point(15, 580)
    $GroupBoxCloud.Size = New-Object System.Drawing.Size(550, 160)

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
    $TextCloudToken.Size = New-Object System.Drawing.Size(520, 25)
    $TextCloudToken.UseSystemPasswordChar = $true
    $TextCloudToken.Enabled = $CheckBoxUseCloud.Checked
    if ($editBase) {
        $tokenKey = "$($editBase.Tag)__YADiskToken"
        if ($Script:AllSecrets.ContainsKey($tokenKey)) {
            $TextCloudToken.Text = [string]$Script:AllSecrets[$tokenKey]
        }
    }

    $LabelCloudKeep = New-Object System.Windows.Forms.Label
    $LabelCloudKeep.Text = "Хранить копий в облаке:"
    $LabelCloudKeep.Location = New-Object System.Drawing.Point(10, 105)
    $LabelCloudKeep.Size = New-Object System.Drawing.Size(200, 20)
    $LabelCloudKeep.Enabled = $CheckBoxUseCloud.Checked

    $NumericCloudKeep = New-Object System.Windows.Forms.NumericUpDown
    $NumericCloudKeep.Location = New-Object System.Drawing.Point(220, 105)
    $NumericCloudKeep.Size = New-Object System.Drawing.Size(100, 25)
    $NumericCloudKeep.Minimum = 0
    $NumericCloudKeep.Maximum = 999
    $NumericCloudKeep.Value = if ($editBase -and $editBase.PSObject.Properties.Name -contains 'CloudKeep') { $editBase.CloudKeep } else { 3 }
    $NumericCloudKeep.Enabled = $CheckBoxUseCloud.Checked

    $ButtonTestCloud = New-Object System.Windows.Forms.Button
    $ButtonTestCloud.Text = "Тест подключения"
    $ButtonTestCloud.Location = New-Object System.Drawing.Point(340, 105)
    $ButtonTestCloud.Size = New-Object System.Drawing.Size(140, 25)
    $ButtonTestCloud.BackColor = [System.Drawing.Color]::LightBlue
    $ButtonTestCloud.Enabled = $CheckBoxUseCloud.Checked

    $GroupBoxCloud.Controls.AddRange(@($CheckBoxUseCloud, $LabelCloudToken, $TextCloudToken, $LabelCloudKeep, $NumericCloudKeep, $ButtonTestCloud))

    # Кнопки
    $ButtonOK = New-Object System.Windows.Forms.Button
    $ButtonOK.Text = "OK"
    $ButtonOK.Location = New-Object System.Drawing.Point(400, 750)
    $ButtonOK.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $ButtonCancel = New-Object System.Windows.Forms.Button
    $ButtonCancel.Text = "Отмена"
    $ButtonCancel.Location = New-Object System.Drawing.Point(490, 750)
    $ButtonCancel.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $BaseDialog.Controls.AddRange(@($LabelName, $TextName, $LabelType, $ComboType, $LabelSource, $TextSource, $ButtonBrowseSource, $LabelDest, $TextDest, $ButtonBrowseDest, $LabelKeep, $NumericKeep, $GroupBoxDT, $GroupBoxCloud, $ButtonOK, $ButtonCancel))

    # Обработчики событий
    $ComboType.Add_SelectedIndexChanged({
        try {
            $GroupBoxDT.Enabled = ($ComboType.Text -eq "DT")
        }
        catch {
            # Игнорируем ошибки при инициализации
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

        # Дополнительная валидация для DT
        if ($ComboType.Text -eq "DT") {
            if ([string]::IsNullOrWhiteSpace($TextExe.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Для выгрузки DT обязательно указать путь к 1cestart.exe!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if (-not (Test-Path $TextExe.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Указанный файл 1cestart.exe не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        }

        # Проверка уникальности имени (кроме редактирования той же базы)
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

        if ($ComboType.Text -eq "DT") {
            if (-not [string]::IsNullOrWhiteSpace($TextExe.Text)) {
                $baseConfig.ExePath = $TextExe.Text.Trim()
            }

            # Определение служб для остановки
            if ($CheckBoxAutoServices.Checked) {
                $webServices = Get-WebServices
                if ($webServices -and ($webServices.Name -contains 'Apache2.4')) {
                    $baseConfig.StopServices = @('Apache2.4')
                } elseif ($webServices) {
                    $baseConfig.StopServices = $webServices | Select-Object -ExpandProperty Name
                } else {
                    $baseConfig.StopServices = @()
                }
            } elseif ($CheckBoxManualServices.Checked) {
                # Показать диалог выбора служб
                $selectedServices = Show-ServicesSelectionDialog
                $baseConfig.StopServices = if ($selectedServices) { $selectedServices } else { @() }
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

        Update-BasesListView
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
        $CheckedListBox.Items.Add($displayText) | Out-Null
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

# === ШАГ 2: НАСТРОЙКА TELEGRAM ===
function Show-Step2 {
    $ContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Настройка уведомлений Telegram"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(900, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Включение Telegram
    $Script:CheckBoxTelegramEnabled = New-Object System.Windows.Forms.CheckBox
    $Script:CheckBoxTelegramEnabled.Text = "Включить отправку отчётов в Telegram"
    $Script:CheckBoxTelegramEnabled.Location = New-Object System.Drawing.Point(20, 70)
    $Script:CheckBoxTelegramEnabled.Size = New-Object System.Drawing.Size(400, 25)
    $Script:CheckBoxTelegramEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # Основной бот
    $Script:GroupBoxMain = New-Object System.Windows.Forms.GroupBox
    $Script:GroupBoxMain.Text = "Основной бот"
    $Script:GroupBoxMain.Location = New-Object System.Drawing.Point(20, 110)
    $Script:GroupBoxMain.Size = New-Object System.Drawing.Size(900, 120)

    $LabelToken = New-Object System.Windows.Forms.Label
    $LabelToken.Text = "Токен бота:"
    $LabelToken.Location = New-Object System.Drawing.Point(15, 25)
    $LabelToken.Size = New-Object System.Drawing.Size(100, 20)

    $Script:TextTelegramToken = New-Object System.Windows.Forms.TextBox
    $Script:TextTelegramToken.Location = New-Object System.Drawing.Point(15, 50)
    $Script:TextTelegramToken.Size = New-Object System.Drawing.Size(400, 25)
    $Script:TextTelegramToken.UseSystemPasswordChar = $true

    $LabelChatId = New-Object System.Windows.Forms.Label
    $LabelChatId.Text = "ID чата:"
    $LabelChatId.Location = New-Object System.Drawing.Point(450, 25)
    $LabelChatId.Size = New-Object System.Drawing.Size(100, 20)

    $Script:TextTelegramChatId = New-Object System.Windows.Forms.TextBox
    $Script:TextTelegramChatId.Location = New-Object System.Drawing.Point(450, 50)
    $Script:TextTelegramChatId.Size = New-Object System.Drawing.Size(200, 25)

    $Script:CheckBoxOnlyErrors = New-Object System.Windows.Forms.CheckBox
    $Script:CheckBoxOnlyErrors.Text = "Отправлять только при ошибках"
    $Script:CheckBoxOnlyErrors.Location = New-Object System.Drawing.Point(15, 85)
    $Script:CheckBoxOnlyErrors.Size = New-Object System.Drawing.Size(300, 25)

    $Script:GroupBoxMain.Controls.AddRange(@($LabelToken, $Script:TextTelegramToken, $LabelChatId, $Script:TextTelegramChatId, $Script:CheckBoxOnlyErrors))

    # Второй бот
    $Script:GroupBoxSecondary = New-Object System.Windows.Forms.GroupBox
    $Script:GroupBoxSecondary.Text = "Дополнительный бот для ответственного"
    $Script:GroupBoxSecondary.Location = New-Object System.Drawing.Point(20, 250)
    $Script:GroupBoxSecondary.Size = New-Object System.Drawing.Size(900, 120)

    $Script:CheckBoxSecondaryBot = New-Object System.Windows.Forms.CheckBox
    $Script:CheckBoxSecondaryBot.Text = "Включить дополнительный бот"
    $Script:CheckBoxSecondaryBot.Location = New-Object System.Drawing.Point(15, 25)
    $Script:CheckBoxSecondaryBot.Size = New-Object System.Drawing.Size(300, 25)

    $LabelSecondaryToken = New-Object System.Windows.Forms.Label
    $LabelSecondaryToken.Text = "Токен второго бота:"
    $LabelSecondaryToken.Location = New-Object System.Drawing.Point(15, 55)
    $LabelSecondaryToken.Size = New-Object System.Drawing.Size(120, 20)

    $Script:TextSecondaryToken = New-Object System.Windows.Forms.TextBox
    $Script:TextSecondaryToken.Location = New-Object System.Drawing.Point(15, 80)
    $Script:TextSecondaryToken.Size = New-Object System.Drawing.Size(400, 25)
    $Script:TextSecondaryToken.UseSystemPasswordChar = $true
    $Script:TextSecondaryToken.Enabled = $false

    $LabelSecondaryChatId = New-Object System.Windows.Forms.Label
    $LabelSecondaryChatId.Text = "ID чата второго бота:"
    $LabelSecondaryChatId.Location = New-Object System.Drawing.Point(450, 55)
    $LabelSecondaryChatId.Size = New-Object System.Drawing.Size(140, 20)

    $Script:TextSecondaryChatId = New-Object System.Windows.Forms.TextBox
    $Script:TextSecondaryChatId.Location = New-Object System.Drawing.Point(450, 80)
    $Script:TextSecondaryChatId.Size = New-Object System.Drawing.Size(200, 25)
    $Script:TextSecondaryChatId.Enabled = $false

    $Script:GroupBoxSecondary.Controls.AddRange(@($Script:CheckBoxSecondaryBot, $LabelSecondaryToken, $Script:TextSecondaryToken, $LabelSecondaryChatId, $Script:TextSecondaryChatId))

    # Тест
    $Script:ButtonTestTelegram = New-Object System.Windows.Forms.Button
    $Script:ButtonTestTelegram.Text = "Отправить тестовое сообщение"
    $Script:ButtonTestTelegram.Location = New-Object System.Drawing.Point(20, 390)
    $Script:ButtonTestTelegram.Size = New-Object System.Drawing.Size(200, 35)
    $Script:ButtonTestTelegram.BackColor = [System.Drawing.Color]::LightBlue

    # Информация
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Text = "Настройте Telegram-ботов для получения отчётов о резервном копировании.`nЧтобы создать бота, напишите @BotFather в Telegram и следуйте инструкциям.`nДля получения ID чата напишите @userinfobot в Telegram."
    $InfoLabel.Location = New-Object System.Drawing.Point(20, 440)
    $InfoLabel.Size = New-Object System.Drawing.Size(900, 60)
    $InfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $InfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen

    $ContentPanel.Controls.AddRange(@($TitleLabel, $Script:CheckBoxTelegramEnabled, $Script:GroupBoxMain, $Script:GroupBoxSecondary, $Script:ButtonTestTelegram, $InfoLabel))

    # Обработчики событий
    $Script:CheckBoxTelegramEnabled.Add_CheckedChanged({
        try {
            $enabled = $Script:CheckBoxTelegramEnabled.Checked
            $Script:GroupBoxMain.Enabled = $enabled
            $Script:GroupBoxSecondary.Enabled = $enabled
            $Script:ButtonTestTelegram.Enabled = $enabled
            if (-not $enabled -and $Script:CheckBoxSecondaryBot) {
                $Script:CheckBoxSecondaryBot.Checked = $false
            }
        }
        catch {
            # Игнорируем ошибки при инициализации
        }
    })

    $Script:CheckBoxSecondaryBot.Add_CheckedChanged({
        try {
            $enabled = $Script:CheckBoxSecondaryBot.Checked
            if ($Script:TextSecondaryToken) { $Script:TextSecondaryToken.Enabled = $enabled }
            if ($Script:TextSecondaryChatId) { $Script:TextSecondaryChatId.Enabled = $enabled }
        }
        catch {
            # Игнорируем ошибки при инициализации
        }
    })

    $ButtonTestTelegram.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($Script:TextTelegramToken.Text) -or [string]::IsNullOrWhiteSpace($Script:TextTelegramChatId.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Заполните токен и ID чата основного бота!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            Test-TelegramConnection
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Загрузка существующих настроек
    $settingsData = Load-ExistingSettings
    if ($settingsData.ContainsKey('Telegram')) {
        $telegramData = ConvertTo-Hashtable $settingsData['Telegram']
        if ($telegramData.ContainsKey('Enabled')) { $Script:CheckBoxTelegramEnabled.Checked = [bool]$telegramData['Enabled'] }
        if ($telegramData.ContainsKey('ChatId')) { $Script:TextTelegramChatId.Text = [string]$telegramData['ChatId'] }
        if ($telegramData.ContainsKey('NotifyOnlyOnErrors')) { $Script:CheckBoxOnlyErrors.Checked = [bool]$telegramData['NotifyOnlyOnErrors'] }
        if ($telegramData.ContainsKey('SecondaryChatId') -and $telegramData['SecondaryChatId']) {
            $Script:TextSecondaryChatId.Text = [string]$telegramData['SecondaryChatId']
            $Script:CheckBoxSecondaryBot.Checked = $true
        }
    }

    # Загрузка токенов из секретов
    if ($Script:AllSecrets.ContainsKey('__TELEGRAM_BOT_TOKEN')) {
        $Script:TextTelegramToken.Text = [string]$Script:AllSecrets['__TELEGRAM_BOT_TOKEN']
    }
    if ($Script:AllSecrets.ContainsKey('__TELEGRAM_SECONDARY_BOT_TOKEN')) {
        $Script:TextSecondaryToken.Text = [string]$Script:AllSecrets['__TELEGRAM_SECONDARY_BOT_TOKEN']
    }

    # Первоначальная настройка состояния
    try {
        $enabled = $Script:CheckBoxTelegramEnabled.Checked
        $Script:GroupBoxMain.Enabled = $enabled
        $Script:GroupBoxSecondary.Enabled = $enabled
        $Script:ButtonTestTelegram.Enabled = $enabled

        $secondaryEnabled = $Script:CheckBoxSecondaryBot.Checked
        $Script:TextSecondaryToken.Enabled = $secondaryEnabled
        $Script:TextSecondaryChatId.Enabled = $secondaryEnabled
    }
    catch {
        # Игнорируем ошибки при инициализации
    }
}

function Test-YandexDiskConnection {
    param(
        [Parameter(Mandatory)][string]$Token,
        [string]$BaseName = 'Test'
    )

    try {
        # Проверяем наличие модуля
        $yandexModulePath = Join-Path $PSScriptRoot '..\modules\Cloud.YandexDisk.psm1'
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
    try {
        $telegramModulePath = Join-Path $PSScriptRoot '..\modules\Notifications.Telegram.psm1'
        if (-not (Test-Path $telegramModulePath)) {
            [System.Windows.Forms.MessageBox]::Show("Модуль Notifications.Telegram.psm1 не найден!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        Import-Module -Force -DisableNameChecking $telegramModulePath -ErrorAction Stop

        $testMessage = "Тест уведомлений от системы резервного копирования 1С`n" +
                      "Дата и время: {0}`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') +
                      "Статус: Настройка успешно завершена"

        # Тест основного бота
        Send-TelegramMessage -Token $Script:TextTelegramToken.Text -ChatId $Script:TextTelegramChatId.Text -Text $testMessage
        [System.Windows.Forms.MessageBox]::Show("Тестовое сообщение успешно отправлено в основной чат!", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Тест второго бота
        if ($Script:CheckBoxSecondaryBot.Checked -and -not [string]::IsNullOrWhiteSpace($Script:TextSecondaryToken.Text) -and -not [string]::IsNullOrWhiteSpace($Script:TextSecondaryChatId.Text)) {
            try {
                Send-TelegramMessage -Token $Script:TextSecondaryToken.Text -ChatId $Script:TextSecondaryChatId.Text -Text $testMessage
                [System.Windows.Forms.MessageBox]::Show("Тестовое сообщение также успешно отправлено во второй чат!", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Основной бот работает, но не удалось отправить сообщение во второй чат:`n$($_.Exception.Message)", "Частичный успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Не удалось отправить тестовое сообщение:`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# === ШАГ 3: ФИНАЛЬНЫЕ НАСТРОЙКИ ===
function Show-Step3 {
    $ContentPanel.Controls.Clear()

    # Заголовок
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Завершение настройки"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TitleLabel.Size = New-Object System.Drawing.Size(900, 30)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue

    # Действие после бэкапа
    $GroupBoxAfter = New-Object System.Windows.Forms.GroupBox
    $GroupBoxAfter.Text = "Действие после завершения резервного копирования"
    $GroupBoxAfter.Location = New-Object System.Drawing.Point(20, 70)
    $GroupBoxAfter.Size = New-Object System.Drawing.Size(900, 120)

    $Script:RadioShutdown = New-Object System.Windows.Forms.RadioButton
    $Script:RadioShutdown.Text = "Выключить компьютер"
    $Script:RadioShutdown.Location = New-Object System.Drawing.Point(15, 30)
    $Script:RadioShutdown.Size = New-Object System.Drawing.Size(250, 25)

    $Script:RadioRestart = New-Object System.Windows.Forms.RadioButton
    $Script:RadioRestart.Text = "Перезагрузить компьютер"
    $Script:RadioRestart.Location = New-Object System.Drawing.Point(15, 60)
    $Script:RadioRestart.Size = New-Object System.Drawing.Size(250, 25)

    $Script:RadioNothing = New-Object System.Windows.Forms.RadioButton
    $Script:RadioNothing.Text = "Ничего не делать"
    $Script:RadioNothing.Location = New-Object System.Drawing.Point(15, 90)
    $Script:RadioNothing.Size = New-Object System.Drawing.Size(250, 25)
    $Script:RadioNothing.Checked = $true

    $GroupBoxAfter.Controls.AddRange(@($Script:RadioShutdown, $Script:RadioRestart, $Script:RadioNothing))

    # Сводка настроек
    $GroupBoxSummary = New-Object System.Windows.Forms.GroupBox
    $GroupBoxSummary.Text = "Сводка настроек"
    $GroupBoxSummary.Location = New-Object System.Drawing.Point(20, 210)
    $GroupBoxSummary.Size = New-Object System.Drawing.Size(900, 280)

    $Script:TextSummary = New-Object System.Windows.Forms.TextBox
    $Script:TextSummary.Location = New-Object System.Drawing.Point(15, 25)
    $Script:TextSummary.Size = New-Object System.Drawing.Size(870, 240)
    $Script:TextSummary.Multiline = $true
    $Script:TextSummary.ScrollBars = "Vertical"
    $Script:TextSummary.ReadOnly = $true
    $Script:TextSummary.Font = New-Object System.Drawing.Font("Consolas", 9)

    $GroupBoxSummary.Controls.Add($Script:TextSummary)

    $ContentPanel.Controls.AddRange(@($TitleLabel, $GroupBoxAfter, $GroupBoxSummary))

    # Загрузка существующих настроек
    $settingsData = Load-ExistingSettings
    if ($settingsData.ContainsKey('AfterBackup')) {
        switch ([int]$settingsData['AfterBackup']) {
            1 { $Script:RadioShutdown.Checked = $true }
            2 { $Script:RadioRestart.Checked = $true }
            default { $Script:RadioNothing.Checked = $true }
        }
    }

    Update-Summary
}

function Update-Summary {
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
            $summary += "    Копий: $($base.Keep), Облако: $(if ($base.CloudType -eq 'Yandex.Disk') { 'Да' } else { 'Нет' })"
            if ($base.BackupType -eq 'DT' -and $base.StopServices) {
                $summary += "    Остановка служб: $($base.StopServices -join ', ')"
            }
            $summary += ""
        }
    }

    # Telegram
    $summary += "УВЕДОМЛЕНИЯ TELEGRAM:"
    if ($Script:CheckBoxTelegramEnabled -and $Script:CheckBoxTelegramEnabled.Checked) {
        $summary += "  Включено"
        $summary += "  Chat ID: $($Script:TextTelegramChatId.Text)"
        $summary += "  Только ошибки: $(if ($Script:CheckBoxOnlyErrors.Checked) { 'Да' } else { 'Нет' })"
        if ($Script:CheckBoxSecondaryBot -and $Script:CheckBoxSecondaryBot.Checked) {
            $summary += "  Второй бот: Да (Chat ID: $($Script:TextSecondaryChatId.Text))"
        }
    } else {
        $summary += "  Отключено"
    }
    $summary += ""

    # Действие после бэкапа
    $summary += "ДЕЙСТВИЕ ПОСЛЕ РЕЗЕРВНОГО КОПИРОВАНИЯ:"
    if ($Script:RadioShutdown -and $Script:RadioShutdown.Checked) {
        $summary += "  Выключение компьютера"
    } elseif ($Script:RadioRestart -and $Script:RadioRestart.Checked) {
        $summary += "  Перезагрузка компьютера"
    } else {
        $summary += "  Без действий"
    }

    if ($Script:TextSummary) {
        $Script:TextSummary.Text = $summary -join "`r`n"
    }
}

# === НАВИГАЦИЯ ===
function Update-Progress {
    $ProgressBar.Value = $Script:CurrentStep
    switch ($Script:CurrentStep) {
        1 { $ProgressLabel.Text = "Шаг 1 из 3: Настройка баз данных" }
        2 { $ProgressLabel.Text = "Шаг 2 из 3: Настройка Telegram уведомлений" }
        3 { $ProgressLabel.Text = "Шаг 3 из 3: Завершение настройки" }
    }

    $ButtonBack.Enabled = ($Script:CurrentStep -gt 1)
    $ButtonNext.Visible = ($Script:CurrentStep -lt 3)
    $ButtonFinish.Visible = ($Script:CurrentStep -eq 3)

    switch ($Script:CurrentStep) {
        1 { Show-Step1 }
        2 { Show-Step2 }
        3 { Show-Step3 }
    }
}

$ButtonNext.Add_Click({
    if ($Script:CurrentStep -eq 1) {
        if ($Script:AllBases.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Настройте хотя бы одну базу данных!", "Внимание", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }

    if ($Script:CurrentStep -lt $Script:MaxSteps) {
        $Script:CurrentStep++
        Update-Progress
    }
})

$ButtonBack.Add_Click({
    if ($Script:CurrentStep -gt 1) {
        $Script:CurrentStep--
        Update-Progress
    }
})

$ButtonFinish.Add_Click({
    # Сохранение всех настроек
    try {
        Save-AllSettings
        [System.Windows.Forms.MessageBox]::Show("Настройка успешно завершена!`n`nТеперь вы можете запускать резервное копирование через Run-Backup.cmd", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $MainForm.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при сохранении настроек:`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

function Save-AllSettings {
    # Сохранение конфигураций баз
    foreach ($base in $Script:AllBases) {
        $configPath = Join-Path $Script:BasesDir "$($base.Tag).json"
        $base | ConvertTo-Json -Depth 8 | Set-Content -Path $configPath -Encoding UTF8
    }

    # Подготовка секретов
    if ($Script:CheckBoxTelegramEnabled.Checked -and -not [string]::IsNullOrWhiteSpace($Script:TextTelegramToken.Text)) {
        $Script:AllSecrets['__TELEGRAM_BOT_TOKEN'] = $Script:TextTelegramToken.Text
    } elseif ($Script:AllSecrets.ContainsKey('__TELEGRAM_BOT_TOKEN')) {
        $Script:AllSecrets.Remove('__TELEGRAM_BOT_TOKEN')
    }

    if ($Script:CheckBoxSecondaryBot.Checked -and -not [string]::IsNullOrWhiteSpace($Script:TextSecondaryToken.Text)) {
        $Script:AllSecrets['__TELEGRAM_SECONDARY_BOT_TOKEN'] = $Script:TextSecondaryToken.Text
    } elseif ($Script:AllSecrets.ContainsKey('__TELEGRAM_SECONDARY_BOT_TOKEN')) {
        $Script:AllSecrets.Remove('__TELEGRAM_SECONDARY_BOT_TOKEN')
    }

    # Сохранение секретов
    Encrypt-Secrets -Secrets $Script:AllSecrets -KeyPath $Script:KeyPath -OutFile $Script:SecretsFile

    # Настройки Telegram
    $telegramSettings = @{
        Enabled = if ($Script:CheckBoxTelegramEnabled) { $Script:CheckBoxTelegramEnabled.Checked } else { $false }
        ChatId = if ($Script:TextTelegramChatId) { $Script:TextTelegramChatId.Text.Trim() } else { '' }
        NotifyOnlyOnErrors = if ($Script:CheckBoxOnlyErrors) { $Script:CheckBoxOnlyErrors.Checked } else { $false }
        SecondaryChatId = if ($Script:CheckBoxSecondaryBot.Checked -and $Script:TextSecondaryChatId) { $Script:TextSecondaryChatId.Text.Trim() } else { '' }
    }

    # Действие после бэкапа
    $afterBackup = 3
    if ($Script:RadioShutdown.Checked) { $afterBackup = 1 }
    elseif ($Script:RadioRestart.Checked) { $afterBackup = 2 }

    # Общие настройки
    $settings = @{
        AfterBackup = $afterBackup
        Telegram = $telegramSettings
    }

    $settings | ConvertTo-Json -Depth 6 | Set-Content -Path $Script:SettingsFile -Encoding UTF8
}

# === ЗАГРУЗКА И ЗАПУСК ===
Load-ExistingSettings | Out-Null
Update-Progress

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