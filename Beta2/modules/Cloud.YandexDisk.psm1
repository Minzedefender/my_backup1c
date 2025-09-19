#requires -Version 5.1

function Format-Bytes {
    param([Parameter(Mandatory)][decimal]$Bytes)
    $units = @('b','kb','mb','gb','tb')
    $val = [decimal]$Bytes
    $i = 0
    while ($val -ge 1024 -and $i -lt $units.Count-1) {
        $val = [math]::Round($val/1024,1)
        $i++
    }
    ("{0:N1}{1}" -f $val, $units[$i]).Replace('.',',')
}
function Format-Speed {
    param([Parameter(Mandatory)][double]$BytesPerSec)
    (Format-Bytes ([decimal]$BytesPerSec)) + '/s'
}
function Show-BarLine {
    param(
        [int]$Width = 28,
        [double]$Percent,
        [string]$doneHuman,
        [string]$totalHuman,
        [string]$speedHuman
    )
    if ($Width -lt 10) { $Width = 10 }
    $bars = [int][math]::Round(($Percent/100.0) * $Width)
    $bar = ('#' * $bars).PadRight($Width,' ')
    "{0} {1,3}% {2}/{3} {4}" -f $bar, [int]$Percent, $doneHuman, $totalHuman, $speedHuman
}


function Ensure-YandexDiskFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$RemotePath
    )

    if ([string]::IsNullOrWhiteSpace($RemotePath)) { return }
    $trimmed = $RemotePath.Trim('/')
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return }

    $headers = @{ Authorization = "OAuth $Token" }
    $segments = $trimmed.Split('/')
    $current = ''
    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }
        $current = '/' + (($current.Trim('/') + '/' + $segment).Trim('/'))
        $enc = [Uri]::EscapeDataString($current)
        $uri = "https://cloud-api.yandex.net/v1/disk/resources?path=$enc"
        try {
            Invoke-RestMethod -Uri $uri -Headers $headers -Method Put -ErrorAction Stop | Out-Null
        }
        catch {
            $resp = $_.Exception.Response
            if ($resp -and $resp.StatusCode.value__ -eq 409) {
                # already exists - ignore
            }
            elseif ($resp -and $resp.StatusCode.value__ -eq 401) {
                throw "Yandex Disk authentication failed"
            }
            else {
                throw "Не удалось создать папку '$current' на Яндекс.Диске: $($_.Exception.Message)"
            }
        }
    }
}

function Upload-ToYandexDisk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$RemotePath,
        [int]$ChunkSize = 4MB,
        [int]$BarWidth  = 28
    )

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Файл не найден: $LocalPath"
    }
    $fileInfo = Get-Item -LiteralPath $LocalPath
    $total = [long]$fileInfo.Length
    $done  = 0L

    $headers = @{ Authorization = "OAuth $Token" }

    $parent = Split-Path -Path $RemotePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) { Ensure-YandexDiskFolder -Token $Token -RemotePath $parent }

    $encPath = [Uri]::EscapeDataString($RemotePath)

    # 1) получить URL
    $url = "https://cloud-api.yandex.net/v1/disk/resources/upload?path=$encPath&overwrite=true"
    $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
    if (-not $resp.href) { throw "Не удалось получить URL загрузки" }
    $putUrl = $resp.href

    # 2) PUT потоково
    $req = [System.Net.HttpWebRequest]::Create($putUrl)
    $req.Method = "PUT"
    $req.AllowWriteStreamBuffering = $false
    $req.SendChunked = $true
    $req.Timeout = 10*60*1000
    $req.ReadWriteTimeout = 10*60*1000
    $req.ContentLength = $total

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastTick = 0L

    $fs = [System.IO.File]::OpenRead($LocalPath)
    try {
        $reqStream = $req.GetRequestStream()
        try {
            $buffer = New-Object byte[] $ChunkSize
            while ($true) {
                $read = $fs.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) { break }
                $reqStream.Write($buffer, 0, $read)
                $done += $read

                # прогресс
                $elapsed = [math]::Max(0.2, $sw.Elapsed.TotalSeconds)
                $speed   = ($done - $lastTick) / $elapsed
                $lastTick = $done
                $sw.Restart()

                $pct  = if ($total -gt 0) { 100.0 * $done / $total } else { 0 }
                $line = Show-BarLine -Width $BarWidth -Percent $pct `
                        -doneHuman (Format-Bytes $done) `
                        -totalHuman (Format-Bytes $total) `
                        -speedHuman (Format-Speed $speed)

                try {
                    # корректная перерисовка строки
                    $raw = $Host.UI.RawUI
                    $pos = $raw.CursorPosition
                    $pos.X = 0
                    $raw.CursorPosition = $pos
                    $w = $raw.BufferSize.Width
                    Write-Host ($line.PadRight([math]::Max($w,120))) -NoNewline
                } catch {
                    # запасной путь
                    Write-Host "`r$line" -NoNewline
                }
            }
        } finally { $reqStream.Close() }

        $resp2 = $req.GetResponse()
        $resp2.Close()
        Write-Host ""
    }
    finally {
        $fs.Close()
    }
}

Export-ModuleMember -Function Upload-ToYandexDisk
