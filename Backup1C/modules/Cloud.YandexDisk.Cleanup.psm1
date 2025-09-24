#requires -Version 5.1

function Invoke-YandexDiskRequest {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('Get','Put','Delete')]
        [string]$Method = 'Get',
        $Body = $null
    )

    $headers = @{ Authorization = "OAuth $Token" }
    Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method -Body $Body -ErrorAction Stop
}

function Get-YandexDiskItems {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$RemotePath,
        [int]$Limit = 200
    )

    $items = @()
    $offset = 0
    $encPath = [Uri]::EscapeDataString($RemotePath)

    while ($true) {
        $url = "https://cloud-api.yandex.net/v1/disk/resources?path=$encPath&limit=$Limit&offset=$offset&fields=_embedded.items.name,_embedded.items.type,_embedded.items.created,_embedded.next"
        $resp = Invoke-YandexDiskRequest -Token $Token -Uri $url
        if ($resp._embedded -and $resp._embedded.items) {
            $items += $resp._embedded.items
            if ($resp._embedded.next) {
                $offset += $Limit
                continue
            }
        }
        break
    }
    return $items
}

function Cleanup-YandexDiskFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$RemoteFolder,
        [int]$Keep = 0,
        [string]$FilePrefix = '',
        [int]$Limit = 200,
        [ScriptBlock]$LogAction = $null,
        [switch]$WhatIf
    )

    if ($Keep -le 0) { return @() }

    $invokeLog = {
        param($msg)
        if ($LogAction) {
            & $LogAction $msg
        }
    }

    $items = @()
    try {
        $items = Get-YandexDiskItems -Token $Token -RemotePath $RemoteFolder -Limit $Limit
    }
    catch {
        throw "Не удалось получить список файлов с Я.Диска: $($_.Exception.Message)"
    }

    if (-not $items -or $items.Count -eq 0) { return @() }

    $files = $items | Where-Object {
        $_.type -eq 'file' -and (
            [string]::IsNullOrEmpty($FilePrefix) -or $_.name -like "${FilePrefix}*")
    } | Sort-Object { Get-Date $_.created } -Descending

    if ($files.Count -le $Keep) { return @() }

    $toDelete = @()
    for ($i = $Keep; $i -lt $files.Count; $i++) {
        $toDelete += $files[$i]
    }

    foreach ($item in $toDelete) {
        $path = $RemoteFolder.TrimEnd('/') + '/' + $item.name
        & $invokeLog ("Удаление из облака: {0}" -f $item.name)
        if (-not $WhatIf) {
            $enc = [Uri]::EscapeDataString($path)
            $urlDel = "https://cloud-api.yandex.net/v1/disk/resources?path=$enc&permanently=true"
            try { Invoke-YandexDiskRequest -Token $Token -Uri $urlDel -Method Delete }
            catch {
                & $invokeLog ("[WARN] Не удалось удалить {0}: {1}" -f $item.name, $_.Exception.Message)
            }
        }
    }
    return $toDelete
}

Export-ModuleMember -Function Cleanup-YandexDiskFolder
