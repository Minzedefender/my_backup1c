# core/Pipeline.psm1
#requires -Version 5.1

$ModulesRoot = Join-Path $PSScriptRoot '..\modules'
Import-Module -Force -DisableNameChecking (Join-Path $ModulesRoot 'Common.Crypto.psm1') -ErrorAction Stop
if (Test-Path (Join-Path $ModulesRoot 'Cloud.YandexDisk.psm1')) {
    Import-Module -Force -DisableNameChecking (Join-Path $ModulesRoot 'Cloud.YandexDisk.psm1') -ErrorAction SilentlyContinue
}
if (Test-Path (Join-Path $ModulesRoot 'Cloud.YandexDisk.Cleanup.psm1')) {
    Import-Module -Force -DisableNameChecking (Join-Path $ModulesRoot 'Cloud.YandexDisk.Cleanup.psm1') -ErrorAction SilentlyContinue
}

function ConvertTo-Hashtable {
    param([Parameter(Mandatory)]$Object)
    if ($Object -is [hashtable]) { return $Object }
    $ht = @{}
    $Object.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
    return $ht
}

function Wait-FileStable {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSec = 0,
        [int]$PollSec    = 2,
        [int]$StableTicks= 3,
        [System.Diagnostics.Process]$Process = $null,
        [int]$PostExitWaitSec = 60
    )
    $deadline = if ($TimeoutSec -gt 0) { (Get-Date).AddSeconds($TimeoutSec) } else { $null }
    $exitDeadline = $null
    $last = -1L
    $stable = 0
    while ($true) {
        if (Test-Path -LiteralPath $Path) {
            $len = (Get-Item -LiteralPath $Path).Length
            if ($len -gt 0 -and $len -eq $last) {
                $stable++
                if ($stable -ge $StableTicks) { return $true }
            } else {
                $stable = 0
                $last = $len
            }
        } else {
            $stable = 0
            $last = -1L
        }

        if ($deadline -and (Get-Date) -gt $deadline) { return $false }

        if ($Process) {
            if ($Process.HasExited) {
                if (-not $exitDeadline) { $exitDeadline = (Get-Date).AddSeconds([math]::Max(1,$PostExitWaitSec)) }
                elseif ((Get-Date) -gt $exitDeadline) { return $false }
            }
        }

        Start-Sleep -Seconds $PollSec
    }
}

function Get-ShortPath([string]$Path){
    $sig = @'
using System;
using System.Runtime.InteropServices;
public static class SP {
 [DllImport("kernel32.dll", CharSet=CharSet.Auto)]
 public static extern int GetShortPathName(string l, System.Text.StringBuilder s, int b);
}
'@
    if (-not ([System.Management.Automation.PSTypeName]'SP').Type) { Add-Type $sig -ErrorAction SilentlyContinue | Out-Null }
    try {
        $sb = New-Object System.Text.StringBuilder 260
        [void][SP]::GetShortPathName($Path, $sb, $sb.Capacity)
        $sp = $sb.ToString()
        if ([string]::IsNullOrWhiteSpace($sp)) { return $Path }
        return $sp
    } catch { return $Path }
}

function Find-1CDesignerExe {
    $roots = @('HKLM:\SOFTWARE\1C\1Cv8','HKLM:\SOFTWARE\WOW6432Node\1C\1Cv8')
    foreach ($r in $roots) {
        if (Test-Path $r) {
            foreach ($k in (Get-ChildItem $r -ErrorAction SilentlyContinue)) {
                $bin = (Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue).BinDir
                if ($bin -and (Test-Path $bin)) {
                    $exe = Join-Path $bin '1cv8.exe'
                    if (Test-Path $exe) { return $exe }
                }
            }
        }
    }
    return $null
}

function BuildArgsLine-F([string]$srcFolder,[string]$artifact,[string]$logFile,[string]$login,[string]$password){
    $parts = @('DESIGNER','/F',('"{0}"' -f $srcFolder),'/DisableStartupMessages','/DumpIB',('"{0}"' -f $artifact),'/Out',('"{0}"' -f $logFile))
    if ($login)    { $parts += @('/N',('"{0}"' -f $login)) }
    if ($password) { $parts += @('/P',('"{0}"' -f $password)) }
    ($parts -join ' ')
}
function BuildArgsLine-IBCS([string]$srcFolder,[string]$artifact,[string]$logFile,[string]$login,[string]$password){
    $srcEsc = $srcFolder.Replace('"','""')
    $conn = 'File="{0}";' -f $srcEsc
    $parts = @('DESIGNER','/IBConnectionString',('"{0}"' -f $conn),'/DisableStartupMessages','/DumpIB',('"{0}"' -f $artifact),'/Out',('"{0}"' -f $logFile))
    if ($login)    { $parts += @('/N',('"{0}"' -f $login)) }
    if ($password) { $parts += @('/P',('"{0}"' -f $password)) }
    ($parts -join ' ')
}

function Invoke-Pipeline {
    param([Parameter(Mandatory)][hashtable]$Ctx)

    $log = $Ctx.Log
    $tag = $Ctx.Tag
    & $log ("База: {0}" -f $tag)

    $configFile   = Join-Path $Ctx.ConfigDir  ("{0}.json" -f $tag)
    $configRoot   = $Ctx.ConfigRoot
    $keyPath      = Join-Path $configRoot 'key.bin'
    $secretsPath  = Join-Path $configRoot 'secrets.json.enc'

    if (!(Test-Path $configFile)) { throw "Конфиг базы '$tag' не найден: $configFile" }

    $config  = ConvertTo-Hashtable (Get-Content $configFile -Raw | ConvertFrom-Json)
    if ($config.ContainsKey('Disabled') -and $config['Disabled']) {
        & $log ("База [{0}] отключена (Disabled=true) — пропуск." -f $tag)
        return $null
    }

    $secrets = @{}
    if (Test-Path $secretsPath) {
        try { $secrets = ConvertTo-Hashtable (Decrypt-Secrets -InFile $secretsPath -KeyPath $keyPath) } catch { $secrets = @{} }
    }

    $type       = ("" + $config['BackupType']).ToUpper()
    $src        = $config['SourcePath']
    $dstDir     = $config['DestinationPath']
    $exeCfg     = $config['ExePath']
    $keep       = 0 + ($config['Keep'])
    $useCloud   = ("" + $config['CloudType']) -eq 'Yandex.Disk'
    $cloudKeep  = 0
    if ($config.ContainsKey('CloudKeep')) {
        try { $cloudKeep = [int]$config['CloudKeep'] } catch { $cloudKeep = 0 }
        if ($cloudKeep -lt 0) { $cloudKeep = 0 }
    }

    if ([string]::IsNullOrWhiteSpace($src))    { throw "Не задан SourcePath в конфиге" }
    if ([string]::IsNullOrWhiteSpace($dstDir)) { throw "Не задан DestinationPath в конфиге" }
    if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }

    $ts  = Get-Date -Format 'yyyy_MM_dd_HHmm'
    $ext = if ($type -eq 'DT') { 'dt' } else { '1CD' }
    $artifact = Join-Path $dstDir ("{0}_{1}.{2}" -f $tag, $ts, $ext)

    $stopped = @()
    try {
        if ($type -eq 'DT') {
            # список служб для лога
            $stopSvcsStr = '—'
            if ($config.ContainsKey('StopServices') -and $config['StopServices']) {
                $stopSvcsStr = ($config['StopServices'] -join ', ')
            }
            & $log ("Остановка служб: {0}" -f $stopSvcsStr)

            if ($config.ContainsKey('StopServices') -and $config['StopServices']) {
                foreach ($name in @($config['StopServices'])) {
                    $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
                    if ($null -ne $svc -and $svc.Status -ne 'Stopped') {
                        try { Stop-Service -Name $svc.Name -Force -ErrorAction Stop; $svc.WaitForStatus('Stopped','00:00:30'); $stopped += $svc.Name } catch {}
                    }
                }
            }

            & $log ("Запуск выгрузки в формате DT...")

            $exe = $null
            if ($exeCfg -and (Test-Path $exeCfg)) { $exe = $exeCfg } else { $exe = Find-1CDesignerExe }
            if (-not $exe) { throw "Не найден исполняемый файл 1С (ни ExePath из конфига, ни 1cv8.exe из реестра)" }

            $loginKey = "{0}__DT_Login"    -f $tag
            $passKey  = "{0}__DT_Password" -f $tag
            $login    = if ($secrets.ContainsKey($loginKey)) { [string]$secrets[$loginKey] } else { $null }
            $password = if ($secrets.ContainsKey($passKey))  { [string]$secrets[$passKey] }  else { $null }

            $srcFolder = (Resolve-Path -LiteralPath $src).Path
            $srcFolder = $srcFolder.TrimEnd('\','/')
            $logFile   = Join-Path $dstDir ("{0}_{1}.designer.log" -f $tag,$ts)

            $srcF83 = Get-ShortPath $srcFolder
            $argsF  = BuildArgsLine-F $srcF83 $artifact $logFile $login $password
            $procF  = Start-Process -FilePath $exe -ArgumentList $argsF -WindowStyle Hidden -PassThru
            $ok = Wait-FileStable -Path $artifact -Process $procF
            [void]$procF.WaitForExit()
            if ($procF.ExitCode -ne 0) { & $log ("[WARN] 1cv8 завершился с кодом {0} (режим /F)" -f $procF.ExitCode) }

            if (-not $ok) {
                & $log ("Повторный запуск выгрузки через IBConnectionString")
                $argsIB = BuildArgsLine-IBCS $srcFolder $artifact $logFile $login $password
                $procIB = Start-Process -FilePath $exe -ArgumentList $argsIB -WindowStyle Hidden -PassThru
                $ok = Wait-FileStable -Path $artifact -Process $procIB
                [void]$procIB.WaitForExit()
                if ($procIB.ExitCode -ne 0) { & $log ("[WARN] 1cv8 завершился с кодом {0} (режим IBConnectionString)" -f $procIB.ExitCode) }
                if (-not $ok) {
                    if (Test-Path $logFile) { throw "Не удалось получить стабильный файл .dt. См. лог: $logFile" }
                    else                    { throw "Не удалось получить стабильный файл .dt" }
                }
            }
        }
        elseif ($type -eq '1CD') {
            & $log ("Копирование файла .1CD...")
            Copy-Item -LiteralPath $src -Destination $artifact -Force -ErrorAction Stop
        }
        else {
            throw "Неизвестный тип бэкапа: $type (ожидалось '1CD' или 'DT')"
        }

        & $log ("Файл бэкапа: {0}" -f $artifact)

        # === ЧИСТКА ЛОКАЛЬНО (ТОЛЬКО АРХИВЫ ТЕКУЩЕГО ТИПА) ===
        if ($keep -gt 0) {
            $mask = ("{0}_*.{1}" -f $tag, $ext)
            $all = Get-ChildItem -Path $dstDir -Filter $mask -File | Sort-Object LastWriteTime -Descending
            if ($all.Count -gt $keep) {
                $toDel = $all[$keep..($all.Count-1)]
                foreach ($f in $toDel) {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
                    & $log ("Удалён старый локальный бэкап: {0}" -f $f.Name)
                }
            }
        }

        $cloudResult = if ($useCloud) { 'Не выполнено' } else { 'Не использовалось' }

        $tokenKey = "{0}__YADiskToken" -f $tag
        $token = $null
        $remoteFolder = if ($useCloud) { "/Backups1C/$tag" } else { $null }
        if ($useCloud -and $secrets.ContainsKey($tokenKey) -and $secrets[$tokenKey]) {
            $token = [string]$secrets[$tokenKey]
        }

        # Облако
        if ($useCloud -and (Get-Command Upload-ToYandexDisk -ErrorAction SilentlyContinue)) {
            if ($token) {
                $remoteName   = Split-Path $artifact -Leaf
                & $log ("Выгрузка в облако Я.Диск")
                try {
                    Upload-ToYandexDisk -Token $token -LocalPath $artifact -RemotePath "$remoteFolder/$remoteName" -BarWidth 28
                    & $log ("Выгрузка в облако Я.Диск завершена")
                    $cloudResult = 'Успех'
                }
                catch {
                    $err = $_.Exception.Message
                    if ($err.Length -gt 160) { $err = $err.Substring(0,160).Trim() + '...' }
                    $cloudResult = "Ошибка: $err"
                    & $log ("[WARN] Не удалось отправить на Я.Диск: {0}" -f $_.Exception.Message)
                }
            } else {
                $cloudResult = 'Нет токена'
                & $log ("ВНИМАНИЕ: включено облако, но токен для [{0}] не найден" -f $tag)
            }
        }
        elseif ($useCloud) {
            $cloudResult = 'Команда недоступна'
            & $log ("[WARN] Функция Upload-ToYandexDisk недоступна")
        }

        if ($useCloud -and $cloudKeep -gt 0 -and $token -and $remoteFolder -and (Get-Command Cleanup-YandexDiskFolder -ErrorAction SilentlyContinue)) {
            try {
                $removed = Cleanup-YandexDiskFolder -Token $token -RemoteFolder $remoteFolder -Keep $cloudKeep -FilePrefix "$tag" -LogAction { param($msg) & $log $msg }
                if ($removed -and $removed.Count -gt 0) {
                    $names = ($removed | ForEach-Object { $_.name }) -join ', '
                    & $log ("Удалены старые файлы из облака: {0}" -f $names)
                }
                else {
                    & $log ("Файлы в облаке актуальны, чистка не требовалась")
                }
            }
            catch {
                & $log ("[WARN] Ошибка очистки Я.Диска: {0}" -f $_.Exception.Message)
            }
        }

        & $log ("=== Успешно завершено [{0}] ===" -f $tag)
        return [pscustomobject]@{ Artifact = $artifact; CloudStatus = $cloudResult }

    }
    catch {
        throw ("Сбой бэкапа [{0}]: {1}" -f $tag, $_.Exception.Message)
    }
    finally {
        if ($stopped.Count -gt 0) {
            foreach ($n in $stopped) {
                try {
                    Start-Service -Name $n -ErrorAction Stop
                    (Get-Service -Name $n).WaitForStatus('Running','00:00:30')
                } catch {}
            }
        }
    }
}

Export-ModuleMember -Function Invoke-Pipeline

