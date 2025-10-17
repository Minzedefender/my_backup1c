# modules\System.Scheduler.psm1
# Modul dlya upravleniya raspisaniem rezervnogo kopirovaniya

#requires -Version 5.1

# Funktsiya dlya sozdaniya zadaniya v planirovshhike Windows
function New-BackupScheduledTask {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName,
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [string]$ScriptArguments, # Complete PowerShell arguments
        [Parameter(Mandatory = $true)][string]$Frequency, # Hourly, CustomHours, Daily, Weekly, Monthly
        [string]$Time = "02:00", # HH:MM - not used for Hourly
        [string]$Days, # For Weekly: "Monday,Tuesday,Wednesday"
        [int]$DayOfMonth = 1, # For Monthly
        [int]$Interval = 1, # For CustomHours: interval in hours
        [string]$Description = "Avtomaticheskoe rezervnoe kopirovanie 1C",
        [string]$User = "SYSTEM"
    )

    try {
        # Proveryaem sushhestvovanie zadachi
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        }

        # Sozdaem deistvie - ispolzuem ScriptArguments esli predostavleny
        if ($ScriptArguments) {
            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $ScriptArguments
        } else {
            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`""
        }

        # Sozdaem trigger v zavisimosti ot chastoty
        switch ($Frequency.ToLower()) {
            "hourly" {
                # Every hour trigger - use simple daily trigger (hourly is complex in Task Scheduler)
                $trigger = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 365)
            }
            "customhours" {
                # Custom hour interval trigger
                $trigger = New-ScheduledTaskTrigger -Once -At $Time -RepetitionInterval (New-TimeSpan -Hours $Interval) -RepetitionDuration (New-TimeSpan -Days 365)
            }
            "daily" {
                $trigger = New-ScheduledTaskTrigger -Daily -At $Time
            }
            "weekly" {
                $daysOfWeek = $Days -split ',' | ForEach-Object { $_.Trim() }
                $trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek $daysOfWeek -At $Time
            }
            "monthly" {
                $trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth $DayOfMonth -At $Time
            }
            default {
                throw "Nepodderzhivaemaya chastota: $Frequency"
            }
        }

        # Nastroiki zadachi
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

        # Sozdaem zadachu
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Description $Description

        # Registriruem zadachu
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -User $User -Force

        return @{
            Success = $true
            Message = "Zadacha '$TaskName' uspeshno sozdana"
            TaskName = $TaskName
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Oshibka sozdaniya zadachi: $($_.Exception.Message)"
            Error = $_.Exception
        }
    }
}

# Funktsiya dlya udaleniya zadaniya
function Remove-BackupScheduledTask {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($task) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            return @{
                Success = $true
                Message = "Zadacha '$TaskName' uspeshno udalena"
            }
        } else {
            return @{
                Success = $false
                Message = "Zadacha '$TaskName' ne naidena"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Oshibka udaleniya zadachi: $($_.Exception.Message)"
            Error = $_.Exception
        }
    }
}

# Funktsiya dlya polucheniya informatsii o zadache
function Get-BackupScheduledTaskInfo {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($task) {
            $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName

            return @{
                Success = $true
                Task = $task
                Info = $taskInfo
                State = $task.State
                LastRunTime = $taskInfo.LastRunTime
                NextRunTime = $taskInfo.NextRunTime
                LastTaskResult = $taskInfo.LastTaskResult
            }
        } else {
            return @{
                Success = $false
                Message = "Zadacha '$TaskName' ne naidena"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Oshibka polucheniya informatsii o zadache: $($_.Exception.Message)"
            Error = $_.Exception
        }
    }
}

# Funktsiya dlya zapuska zadachi nemedlenno
function Start-BackupScheduledTask {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    try {
        Start-ScheduledTask -TaskName $TaskName
        return @{
            Success = $true
            Message = "Zadacha '$TaskName' zapushhena"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Oshibka zapuska zadachi: $($_.Exception.Message)"
            Error = $_.Exception
        }
    }
}

# Funktsiya dlya polucheniya vsekh zadach rezervnogo kopirovaniya
function Get-AllBackupScheduledTasks {
    try {
        $allTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*Backup1C*" -or $_.Description -like "*rezervn*kopir*" }

        $taskList = @()
        foreach ($task in $allTasks) {
            $info = Get-ScheduledTaskInfo -TaskName $task.TaskName
            $taskList += @{
                Name = $task.TaskName
                State = $task.State
                Description = $task.Description
                LastRunTime = $info.LastRunTime
                NextRunTime = $info.NextRunTime
                LastResult = $info.LastTaskResult
            }
        }

        return @{
            Success = $true
            Tasks = $taskList
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Oshibka polucheniya spiska zadach: $($_.Exception.Message)"
            Error = $_.Exception
        }
    }
}

# Funktsiya dlya preobrazovaniya vremeni v stroku
function Format-ScheduleTime {
    param([DateTime]$Time)
    return $Time.ToString("HH:mm")
}

# Funktsiya dlya preobrazovaniya dnei nedeli v stroku
function Format-ScheduleDays {
    param([string[]]$Days)
    $dayTranslation = @{
        'Monday' = 'Ponedelnik'
        'Tuesday' = 'Vtornik'
        'Wednesday' = 'Sreda'
        'Thursday' = 'Chetverg'
        'Friday' = 'Pyatnitsa'
        'Saturday' = 'Subbota'
        'Sunday' = 'Voskresene'
    }

    $translatedDays = $Days | ForEach-Object { $dayTranslation[$_] }
    return $translatedDays -join ", "
}

# Export funktsii
Export-ModuleMember -Function New-BackupScheduledTask, Remove-BackupScheduledTask, Get-BackupScheduledTaskInfo, Start-BackupScheduledTask, Get-AllBackupScheduledTasks, Format-ScheduleTime, Format-ScheduleDays