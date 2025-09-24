#requires -Version 5.1
# Графический редактор конфигураций баз 1С

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Пути
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$ConfigRoot = Join-Path $ProjectRoot 'config'
$BasesDir = Join-Path $ConfigRoot 'bases'

if (-not (Test-Path $ConfigRoot)) { New-Item -ItemType Directory -Path $ConfigRoot -Force | Out-Null }
if (-not (Test-Path $BasesDir)) { New-Item -ItemType Directory -Path $BasesDir -Force | Out-Null }

# Глобальные переменные
$Script:CurrentBase = $null
$Script:HasChanges = $false

# === ГЛАВНАЯ ФОРМА ===
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Редактор конфигураций резервного копирования 1С"
$Form.Size = New-Object System.Drawing.Size(900, 700)
$Form.StartPosition = "CenterScreen"
$Form.MinimumSize = New-Object System.Drawing.Size(900, 700)
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# === ЛЕВАЯ ПАНЕЛЬ СО СПИСКОМ БАЗ ===
$GroupBoxList = New-Object System.Windows.Forms.GroupBox
$GroupBoxList.Text = "Базы данных"
$GroupBoxList.Location = New-Object System.Drawing.Point(12, 12)
$GroupBoxList.Size = New-Object System.Drawing.Size(250, 600)
$GroupBoxList.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$ListBoxBases = New-Object System.Windows.Forms.ListBox
$ListBoxBases.Location = New-Object System.Drawing.Point(10, 25)
$ListBoxBases.Size = New-Object System.Drawing.Size(230, 520)
$ListBoxBases.Font = New-Object System.Drawing.Font("Consolas", 10)
$ListBoxBases.IntegralHeight = $false

$ButtonRefresh = New-Object System.Windows.Forms.Button
$ButtonRefresh.Text = "Обновить"
$ButtonRefresh.Location = New-Object System.Drawing.Point(10, 555)
$ButtonRefresh.Size = New-Object System.Drawing.Size(110, 35)
$ButtonRefresh.UseVisualStyleBackColor = $true

$ButtonDelete = New-Object System.Windows.Forms.Button
$ButtonDelete.Text = "Удалить"
$ButtonDelete.Location = New-Object System.Drawing.Point(130, 555)
$ButtonDelete.Size = New-Object System.Drawing.Size(110, 35)
$ButtonDelete.UseVisualStyleBackColor = $true
$ButtonDelete.Enabled = $false

$GroupBoxList.Controls.AddRange(@($ListBoxBases, $ButtonRefresh, $ButtonDelete))

# === ПРАВАЯ ПАНЕЛЬ С ПАРАМЕТРАМИ ===
$GroupBoxParams = New-Object System.Windows.Forms.GroupBox
$GroupBoxParams.Text = "Параметры базы"
$GroupBoxParams.Location = New-Object System.Drawing.Point(275, 12)
$GroupBoxParams.Size = New-Object System.Drawing.Size(600, 600)
$GroupBoxParams.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$GroupBoxParams.Enabled = $false

# --- Тип бэкапа (строка 1) ---
$LabelType = New-Object System.Windows.Forms.Label
$LabelType.Text = "Тип резервной копии:"
$LabelType.Location = New-Object System.Drawing.Point(15, 30)
$LabelType.Size = New-Object System.Drawing.Size(200, 20)
$LabelType.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$ComboType = New-Object System.Windows.Forms.ComboBox
$ComboType.Items.AddRange(@("1CD", "DT"))
$ComboType.DropDownStyle = "DropDownList"
$ComboType.Location = New-Object System.Drawing.Point(15, 55)
$ComboType.Size = New-Object System.Drawing.Size(570, 25)
$ComboType.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Путь к базе (строка 2) ---
$LabelSource = New-Object System.Windows.Forms.Label
$LabelSource.Text = "Путь к базе данных:"
$LabelSource.Location = New-Object System.Drawing.Point(15, 95)
$LabelSource.Size = New-Object System.Drawing.Size(200, 20)
$LabelSource.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$TextSource = New-Object System.Windows.Forms.TextBox
$TextSource.Location = New-Object System.Drawing.Point(15, 120)
$TextSource.Size = New-Object System.Drawing.Size(480, 25)
$TextSource.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$ButtonBrowseSource = New-Object System.Windows.Forms.Button
$ButtonBrowseSource.Text = "Обзор..."
$ButtonBrowseSource.Location = New-Object System.Drawing.Point(505, 119)
$ButtonBrowseSource.Size = New-Object System.Drawing.Size(80, 27)
$ButtonBrowseSource.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ButtonBrowseSource.UseVisualStyleBackColor = $true

# --- Папка для бэкапов (строка 3) ---
$LabelDest = New-Object System.Windows.Forms.Label
$LabelDest.Text = "Папка для резервных копий:"
$LabelDest.Location = New-Object System.Drawing.Point(15, 160)
$LabelDest.Size = New-Object System.Drawing.Size(250, 20)
$LabelDest.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$TextDest = New-Object System.Windows.Forms.TextBox
$TextDest.Location = New-Object System.Drawing.Point(15, 185)
$TextDest.Size = New-Object System.Drawing.Size(480, 25)
$TextDest.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$ButtonBrowseDest = New-Object System.Windows.Forms.Button
$ButtonBrowseDest.Text = "Обзор..."
$ButtonBrowseDest.Location = New-Object System.Drawing.Point(505, 184)
$ButtonBrowseDest.Size = New-Object System.Drawing.Size(80, 27)
$ButtonBrowseDest.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ButtonBrowseDest.UseVisualStyleBackColor = $true

# --- Количество копий и облако (строка 4) ---
$LabelKeep = New-Object System.Windows.Forms.Label
$LabelKeep.Text = "Хранить копий локально:"
$LabelKeep.Location = New-Object System.Drawing.Point(15, 225)
$LabelKeep.Size = New-Object System.Drawing.Size(170, 20)
$LabelKeep.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$NumericKeep = New-Object System.Windows.Forms.NumericUpDown
$NumericKeep.Location = New-Object System.Drawing.Point(15, 250)
$NumericKeep.Size = New-Object System.Drawing.Size(100, 25)
$NumericKeep.Minimum = 1
$NumericKeep.Maximum = 999
$NumericKeep.Value = 3
$NumericKeep.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$LabelCloud = New-Object System.Windows.Forms.Label
$LabelCloud.Text = "Облачное хранилище:"
$LabelCloud.Location = New-Object System.Drawing.Point(200, 225)
$LabelCloud.Size = New-Object System.Drawing.Size(180, 20)
$LabelCloud.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$ComboCloud = New-Object System.Windows.Forms.ComboBox
$ComboCloud.Items.AddRange(@("Не использовать", "Yandex.Disk"))
$ComboCloud.DropDownStyle = "DropDownList"
$ComboCloud.Location = New-Object System.Drawing.Point(200, 250)
$ComboCloud.Size = New-Object System.Drawing.Size(180, 25)
$ComboCloud.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$LabelCloudKeep = New-Object System.Windows.Forms.Label
$LabelCloudKeep.Text = "Копий в облаке:"
$LabelCloudKeep.Location = New-Object System.Drawing.Point(400, 225)
$LabelCloudKeep.Size = New-Object System.Drawing.Size(120, 20)
$LabelCloudKeep.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$NumericCloudKeep = New-Object System.Windows.Forms.NumericUpDown
$NumericCloudKeep.Location = New-Object System.Drawing.Point(400, 250)
$NumericCloudKeep.Size = New-Object System.Drawing.Size(100, 25)
$NumericCloudKeep.Minimum = 0
$NumericCloudKeep.Maximum = 999
$NumericCloudKeep.Value = 0
$NumericCloudKeep.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Службы для остановки (строка 5, только для DT) ---
$LabelServices = New-Object System.Windows.Forms.Label
$LabelServices.Text = "Службы для остановки (через запятую, только для DT):"
$LabelServices.Location = New-Object System.Drawing.Point(15, 290)
$LabelServices.Size = New-Object System.Drawing.Size(400, 20)
$LabelServices.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$TextServices = New-Object System.Windows.Forms.TextBox
$TextServices.Location = New-Object System.Drawing.Point(15, 315)
$TextServices.Size = New-Object System.Drawing.Size(570, 25)
$TextServices.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$TextServices.Enabled = $false

# --- Путь к 1cv8.exe (строка 6, только для DT) ---
$LabelExe = New-Object System.Windows.Forms.Label
$LabelExe.Text = "Путь к 1cv8.exe (необязательно, только для DT):"
$LabelExe.Location = New-Object System.Drawing.Point(15, 355)
$LabelExe.Size = New-Object System.Drawing.Size(350, 20)
$LabelExe.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$TextExe = New-Object System.Windows.Forms.TextBox
$TextExe.Location = New-Object System.Drawing.Point(15, 380)
$TextExe.Size = New-Object System.Drawing.Size(480, 25)
$TextExe.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$TextExe.Enabled = $false

$ButtonBrowseExe = New-Object System.Windows.Forms.Button
$ButtonBrowseExe.Text = "Обзор..."
$ButtonBrowseExe.Location = New-Object System.Drawing.Point(505, 379)
$ButtonBrowseExe.Size = New-Object System.Drawing.Size(80, 27)
$ButtonBrowseExe.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ButtonBrowseExe.UseVisualStyleBackColor = $true
$ButtonBrowseExe.Enabled = $false

# --- Включение/выключение базы ---
$CheckEnabled = New-Object System.Windows.Forms.CheckBox
$CheckEnabled.Text = "База включена для резервного копирования"
$CheckEnabled.Location = New-Object System.Drawing.Point(15, 430)
$CheckEnabled.Size = New-Object System.Drawing.Size(400, 25)
$CheckEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$CheckEnabled.Checked = $true

# --- Кнопка сохранения ---
$ButtonSave = New-Object System.Windows.Forms.Button
$ButtonSave.Text = "Сохранить изменения"
$ButtonSave.Location = New-Object System.Drawing.Point(15, 550)
$ButtonSave.Size = New-Object System.Drawing.Size(200, 40)
$ButtonSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$ButtonSave.BackColor = [System.Drawing.Color]::LightGreen
$ButtonSave.UseVisualStyleBackColor = $false
$ButtonSave.Enabled = $false

$GroupBoxParams.Controls.AddRange(@(
    $LabelType, $ComboType,
    $LabelSource, $TextSource, $ButtonBrowseSource,
    $LabelDest, $TextDest, $ButtonBrowseDest,
    $LabelKeep, $NumericKeep,
    $LabelCloud, $ComboCloud,
    $LabelCloudKeep, $NumericCloudKeep,
    $LabelServices, $TextServices,
    $LabelExe, $TextExe, $ButtonBrowseExe,
    $CheckEnabled,
    $ButtonSave
))

# Добавляем группы на форму
$Form.Controls.AddRange(@($GroupBoxList, $GroupBoxParams))

# === ФУНКЦИИ ===
function Load-BasesList {
    $ListBoxBases.Items.Clear()
    
    if (-not (Test-Path $BasesDir)) {
        $ListBoxBases.Items.Add("(Нет баз)")
        return
    }
    
    $bases = Get-ChildItem -Path $BasesDir -Filter "*.json" -File -ErrorAction SilentlyContinue
    
    if ($bases.Count -eq 0) {
        $ListBoxBases.Items.Add("(Нет баз)")
        return
    }
    
    foreach ($baseFile in $bases | Sort-Object Name) {
        $baseName = $baseFile.BaseName
        try {
            $config = Get-Content $baseFile.FullName -Raw | ConvertFrom-Json
            $status = if ($config.Disabled) { "[OFF]" } else { "[ON] " }
            $ListBoxBases.Items.Add("$status $baseName")
        } catch {
            $ListBoxBases.Items.Add("[ERR] $baseName")
        }
    }
    
    if ($ListBoxBases.Items.Count -gt 0 -and $ListBoxBases.Items[0] -ne "(Нет баз)") {
        $ListBoxBases.SelectedIndex = 0
    }
}

function Load-BaseConfig {
    param([string]$BaseName)
    
    $Script:CurrentBase = $BaseName
    $configPath = Join-Path $BasesDir "$BaseName.json"
    
    if (-not (Test-Path $configPath)) { return }
    
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        
        # Загружаем значения
        $ComboType.Text = if ($config.BackupType) { $config.BackupType } else { "1CD" }
        $TextSource.Text = if ($config.SourcePath) { $config.SourcePath } else { "" }
        $TextDest.Text = if ($config.DestinationPath) { $config.DestinationPath } else { "" }
        $NumericKeep.Value = if ($config.Keep) { [Math]::Max(1, [int]$config.Keep) } else { 3 }
        
        $ComboCloud.Text = if ($config.CloudType -eq "Yandex.Disk") { "Yandex.Disk" } else { "Не использовать" }
        
        if ($config.PSObject.Properties.Name -contains 'CloudKeep') {
            $NumericCloudKeep.Value = [Math]::Max(0, [int]$config.CloudKeep)
        } else {
            $NumericCloudKeep.Value = 0
        }
        
        if ($config.PSObject.Properties.Name -contains 'StopServices' -and $config.StopServices) {
            $TextServices.Text = ($config.StopServices -join ", ")
        } else {
            $TextServices.Text = ""
        }
        
        if ($config.PSObject.Properties.Name -contains 'ExePath') {
            $TextExe.Text = $config.ExePath
        } else {
            $TextExe.Text = ""
        }
        
        $CheckEnabled.Checked = -not $config.Disabled
        
        # Включаем/отключаем поля в зависимости от типа
        $isDT = ($ComboType.Text -eq "DT")
        $TextServices.Enabled = $isDT
        $TextExe.Enabled = $isDT
        $ButtonBrowseExe.Enabled = $isDT
        
        $GroupBoxParams.Enabled = $true
        $ButtonSave.Enabled = $false
        $ButtonDelete.Enabled = $true
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

function Save-BaseConfig {
    if (-not $Script:CurrentBase) { return }
    
    $configPath = Join-Path $BasesDir "$($Script:CurrentBase).json"
    
    $config = [ordered]@{
        BackupType = $ComboType.Text
        SourcePath = $TextSource.Text
        DestinationPath = $TextDest.Text
        Keep = [int]$NumericKeep.Value
        CloudType = if ($ComboCloud.Text -eq "Yandex.Disk") { "Yandex.Disk" } else { "" }
        CloudKeep = [int]$NumericCloudKeep.Value
        Disabled = -not $CheckEnabled.Checked
    }
    
    if ($ComboType.Text -eq "DT") {
        if ($TextServices.Text.Trim()) {
            $config.StopServices = @($TextServices.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        } else {
            $config.StopServices = @()
        }
        
        if ($TextExe.Text.Trim()) {
            $config.ExePath = $TextExe.Text.Trim()
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
        
        $ButtonSave.Enabled = $false
        $Script:HasChanges = $false
        
        Load-BasesList
        
        # Восстанавливаем выбор
        for ($i = 0; $i -lt $ListBoxBases.Items.Count; $i++) {
            if ($ListBoxBases.Items[$i] -match "$($Script:CurrentBase)$") {
                $ListBoxBases.SelectedIndex = $i
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

# === ОБРАБОТЧИКИ СОБЫТИЙ ===
$ListBoxBases.Add_SelectedIndexChanged({
    if ($ListBoxBases.SelectedItem -and $ListBoxBases.SelectedItem -ne "(Нет баз)") {
        $baseName = $ListBoxBases.SelectedItem -replace '^\[.*?\]\s*', ''
        Load-BaseConfig -BaseName $baseName
    }
})

$ButtonRefresh.Add_Click({ Load-BasesList })

$ButtonDelete.Add_Click({
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
            $GroupBoxParams.Enabled = $false
            $ButtonSave.Enabled = $false
            $ButtonDelete.Enabled = $false
            $Script:CurrentBase = $null
            Load-BasesList
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

$ButtonSave.Add_Click({ Save-BaseConfig })

# Обработчики изменений
$Script:EnableSaveHandler = {
    if ($GroupBoxParams.Enabled) {
        $ButtonSave.Enabled = $true
        $Script:HasChanges = $true
    }
}

$ComboType.Add_SelectedIndexChanged({
    & $Script:EnableSaveHandler
    $isDT = ($ComboType.Text -eq "DT")
    $TextServices.Enabled = $isDT
    $TextExe.Enabled = $isDT
    $ButtonBrowseExe.Enabled = $isDT
})

$TextSource.Add_TextChanged($Script:EnableSaveHandler)
$TextDest.Add_TextChanged($Script:EnableSaveHandler)
$NumericKeep.Add_ValueChanged($Script:EnableSaveHandler)
$ComboCloud.Add_SelectedIndexChanged($Script:EnableSaveHandler)
$NumericCloudKeep.Add_ValueChanged($Script:EnableSaveHandler)
$TextServices.Add_TextChanged($Script:EnableSaveHandler)
$TextExe.Add_TextChanged($Script:EnableSaveHandler)
$CheckEnabled.Add_CheckedChanged($Script:EnableSaveHandler)

# Обработчики кнопок обзора
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
        $dialog.ShowNewFolderButton = $false
        
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
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "1cv8.exe|1cv8.exe|Все exe файлы (*.exe)|*.exe"
    $dialog.Title = "Выберите 1cv8.exe"
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextExe.Text = $dialog.FileName
    }
})

# Обработчик закрытия
$Form.Add_FormClosing({
    param($sender, $e)
    
    if ($Script:HasChanges) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Есть несохраненные изменения. Сохранить перед закрытием?",
            "Несохраненные изменения",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        switch ($result) {
            "Yes" { Save-BaseConfig }
            "Cancel" { $e.Cancel = $true }
        }
    }
})

# === ЗАПУСК ===
Load-BasesList

try {
    [void]$Form.ShowDialog()
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Ошибка: $_",
        "Критическая ошибка",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

$Form.Dispose()