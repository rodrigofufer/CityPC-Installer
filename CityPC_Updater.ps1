[CmdletBinding()]
param(
    [ValidateSet('Check', 'Commit')]
    [string]$Mode = 'Check',
    [string]$ToolId = '',
    [int]$CurrentVersion = 0,
    [string]$TargetPath = '',
    [string]$ResumeToken = '',
    [string]$StatePath = '',
    [int]$ParentPid = 0
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:EngineVersion = 1
$script:ExitContinue = 0
$script:ExitCallerMustClose = 20
$script:ExitResumeFailed = 30
$script:ExitUnsafeState = 31
$script:DownloadDeadlineUtc = $null
$script:ChannelPath = 'preparacion-channel-v2.json'
$script:ManifestName = 'preparacion-bundle-v2.json'
$script:MainBases = @(
    'https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main',
    'https://github.com/rodrigofufer/CityPC-Installer/raw/refs/heads/main'
)
$script:AllowedTools = @{
    instalador = 'Instalador_CityPC.bat'
    anclados   = 'Anclados_y_Limpieza.bat'
    onedrive   = 'Reactivar_OneDrive.bat'
}

function Write-Status {
    param([string]$Level, [string]$Message)
    Write-Host ('  [{0}] {1}' -f $Level.ToUpperInvariant(), $Message)
}

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Test-FileSha256 {
    param([string]$Path, [string]$ExpectedHash)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        return (Get-Sha256 -Path $Path) -ceq $ExpectedHash.ToLowerInvariant()
    } catch {
        return $false
    }
}

function Test-Sha256Text {
    param([string]$Value)
    return ($Value -match '^[0-9a-fA-F]{64}$')
}

function Test-SafeToken {
    param([string]$Value)
    return ($Value -match '^[0-9a-f]{32}$')
}

function Write-JsonAtomic {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$Value)
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $temporary = $Path + '.tmp.' + [Guid]::NewGuid().ToString('N')
    $Value | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Set-StateStatus {
    param([string]$Path, [string]$Status, [string]$Reason = '')
    $state = Read-JsonFile -Path $Path
    $state.status = $Status
    $state.updatedAtUtc = [DateTime]::UtcNow.ToString('o')
    if ($Reason) { $state.reason = $Reason }
    Write-JsonAtomic -Path $Path -Value $state
    return $state
}

function Open-UpdateLock {
    param([Parameter(Mandatory = $true)][string]$StateRoot)
    if (-not (Test-Path -LiteralPath $StateRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $StateRoot -Force | Out-Null
    }
    $lockPath = Join-Path $StateRoot 'update.lock'
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        try {
            return [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        } catch {
            Start-Sleep -Milliseconds 250
        }
    }
    throw 'Otro actualizador CityPC sigue trabajando.'
}

function Invoke-Download {
    param([string]$Url, [string]$Destination, [int]$Attempts = 2)
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if ($null -ne $script:DownloadDeadlineUtc -and [DateTime]::UtcNow -ge $script:DownloadDeadlineUtc) {
            return $false
        }
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        try {
            $cacheBust = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
            $separator = '?'
            if ($Url.Contains('?')) { $separator = '&' }
            Invoke-WebRequest -Uri ($Url + $separator + 'citypc_update=' + $cacheBust) `
                -OutFile $Destination -UseBasicParsing -TimeoutSec 12 -MaximumRedirection 3
            if ((Test-Path -LiteralPath $Destination -PathType Leaf) -and
                ((Get-Item -LiteralPath $Destination).Length -gt 0)) {
                return $true
            }
        } catch {
            Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 1
    }
    return $false
}

function Get-BatchVersion {
    param([Parameter(Mandatory = $true)][string]$Path)
    $content = Get-Content -LiteralPath $Path -Raw -Encoding ASCII
    $match = [Regex]::Match($content, '(?im)^set\s+"LOCAL_VER=([0-9]{1,9})"\s*$')
    if (-not $match.Success) { throw ('LOCAL_VER ausente en ' + [IO.Path]::GetFileName($Path)) }
    return [int]$match.Groups[1].Value
}

function Assert-DownloadedFile {
    param([string]$Path, [string]$Hash, [int64]$MinBytes, [int64]$MaxBytes)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Descarga ausente.' }
    $size = (Get-Item -LiteralPath $Path).Length
    if ($size -lt $MinBytes -or $size -gt $MaxBytes) { throw 'Tamano descargado fuera de rango.' }
    if ((Get-Sha256 -Path $Path) -cne $Hash.ToLowerInvariant()) { throw 'SHA-256 descargado no coincide.' }
}

function Assert-BatchArtifact {
    param([string]$Path, [int]$Version)
    $content = Get-Content -LiteralPath $Path -Raw -Encoding ASCII
    if ($content -notmatch '(?im)^@echo off\s*$') { throw 'Artefacto BAT invalido.' }
    if ((Get-BatchVersion -Path $Path) -ne $Version) { throw 'LOCAL_VER no coincide con manifiesto.' }
    if ($content -notmatch '(?im)CityPC_Updater\.ps1') { throw 'BAT sin actualizador comun.' }
}

function Assert-Channel {
    param($Channel)
    if ([int]$Channel.schema -ne 1) { throw 'Schema de channel no compatible.' }
    if ([string]$Channel.bundleId -cne 'preparacion') { throw 'bundleId de channel invalido.' }
    if ([string]$Channel.tag -notmatch '^preparacion-v[0-9]{1,9}$') { throw 'Tag inmutable invalido.' }
    if ([string]$Channel.manifestFile -cne $script:ManifestName) { throw 'Nombre de manifiesto invalido.' }
    if (-not (Test-Sha256Text -Value ([string]$Channel.manifestSha256))) { throw 'SHA-256 de manifiesto invalido.' }
}

function Assert-Manifest {
    param($Manifest, [string]$ExpectedTag)
    if ([int]$Manifest.schema -ne 3) { throw 'Schema de bundle no compatible.' }
    if ([string]$Manifest.bundleId -cne 'preparacion') { throw 'bundleId invalido.' }
    if ([string]$Manifest.tag -cne $ExpectedTag) { throw 'Tag del bundle no coincide con channel.' }
    if ([string]$Manifest.bundleVersion -notmatch '^[0-9]{1,9}$') { throw 'bundleVersion invalida.' }
    if ([int]$Manifest.bundleVersion -le 0) { throw 'bundleVersion fuera de rango.' }
    if ([int]$Manifest.updater.version -lt 1 -or [string]$Manifest.updater.file -cne 'CityPC_Updater.ps1') {
        throw 'Entrada updater invalida.'
    }
    if (-not (Test-Sha256Text -Value ([string]$Manifest.updater.sha256))) { throw 'SHA updater invalido.' }
    if ([int64]$Manifest.updater.minBytes -lt 1024 -or
        [int64]$Manifest.updater.maxBytes -lt [int64]$Manifest.updater.minBytes -or
        [int64]$Manifest.updater.maxBytes -gt 5242880) { throw 'Limites updater invalidos.' }

    $seenIds = @{}
    $seenFiles = @{}
    $tools = @($Manifest.tools)
    if ($tools.Count -ne $script:AllowedTools.Count) { throw 'El bundle no contiene exactamente tres herramientas.' }
    foreach ($tool in $tools) {
        $id = [string]$tool.id
        if (-not $script:AllowedTools.ContainsKey($id)) { throw ('ToolId no permitido: ' + $id) }
        if ($seenIds.ContainsKey($id)) { throw ('ToolId repetido: ' + $id) }
        if ([string]$tool.file -cne $script:AllowedTools[$id]) { throw ('Archivo no permitido para ' + $id) }
        if ($seenFiles.ContainsKey([string]$tool.file)) { throw 'Archivo repetido en bundle.' }
        if ([string]$tool.version -notmatch '^[0-9]{1,9}$' -or [int]$tool.version -le 0) { throw 'Version de herramienta invalida.' }
        if (-not (Test-Sha256Text -Value ([string]$tool.sha256))) { throw 'SHA de herramienta invalido.' }
        if ($id -ceq 'onedrive') {
            if (-not (Test-Sha256Text -Value ([string]$tool.legacySha256))) {
                throw 'SHA legacy de Reactivar OneDrive invalido.'
            }
        } elseif ([string]$tool.legacySha256) {
            throw 'legacySha256 solo se permite para Reactivar OneDrive.'
        }
        if ([int64]$tool.minBytes -lt 256 -or [int64]$tool.maxBytes -lt [int64]$tool.minBytes -or
            [int64]$tool.maxBytes -gt 5242880) { throw 'Limites de herramienta invalidos.' }
        $seenIds[$id] = $true
        $seenFiles[[string]$tool.file] = $true
    }
}

function Get-RemoteBundle {
    param([string]$ScratchDirectory)
    $channelPath = Join-Path $ScratchDirectory 'channel.json'
    $manifestPath = Join-Path $ScratchDirectory 'manifest.json'
    foreach ($mainBase in $script:MainBases) {
        if (-not (Invoke-Download -Url ($mainBase + '/' + $script:ChannelPath) -Destination $channelPath)) { continue }
        try {
            $channel = Read-JsonFile -Path $channelPath
            Assert-Channel -Channel $channel
            $tag = [string]$channel.tag
            $tagBases = @(
                ('https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/' + $tag),
                ('https://github.com/rodrigofufer/CityPC-Installer/raw/refs/tags/' + $tag)
            )
            foreach ($tagBase in $tagBases) {
                if (-not (Invoke-Download -Url ($tagBase + '/' + $script:ManifestName) -Destination $manifestPath)) { continue }
                if ((Get-Sha256 -Path $manifestPath) -cne ([string]$channel.manifestSha256).ToLowerInvariant()) { continue }
                $manifest = Read-JsonFile -Path $manifestPath
                Assert-Manifest -Manifest $manifest -ExpectedTag $tag
                return [ordered]@{
                    channel = $channel
                    manifest = $manifest
                    manifestSha256 = ([string]$channel.manifestSha256).ToLowerInvariant()
                    tagBase = $tagBase
                }
            }
        } catch {
            Remove-Item -LiteralPath $manifestPath -Force -ErrorAction SilentlyContinue
        }
    }
    return $null
}

function Get-FailureBlock {
    param([string]$Path, [int]$BundleVersion, [string]$ManifestHash)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        $failure = Read-JsonFile -Path $Path
        return (
            [int]$failure.bundleVersion -eq $BundleVersion -and
            [string]$failure.manifestSha256 -ceq $ManifestHash -and
            [DateTime]::Parse([string]$failure.blockedUntilUtc).ToUniversalTime() -gt [DateTime]::UtcNow
        )
    } catch { return $false }
}

function Set-FailureBlock {
    param([string]$Path, [int]$BundleVersion, [string]$ManifestHash, [string]$Reason)
    $failure = [ordered]@{
        bundleVersion = $BundleVersion
        manifestSha256 = $ManifestHash
        reason = $Reason
        createdAtUtc = [DateTime]::UtcNow.ToString('o')
        blockedUntilUtc = [DateTime]::UtcNow.AddMinutes(15).ToString('o')
    }
    Write-JsonAtomic -Path $Path -Value $failure
}

function Wait-ForProcessExit {
    param([int]$ProcessId, [int]$TimeoutSeconds)
    if ($ProcessId -le 0) { return $false }
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

function Get-CallerProcessId {
    $parentProcess = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $PID)
    $parent = [int]$parentProcess.ParentProcessId
    if ($parent -le 0) { throw 'No se pudo identificar el proceso BAT para cierre seguro.' }
    return $parent
}

function Start-Committer {
    param([string]$TransactionStatePath, [int]$CallerProcessId)
    $state = Read-JsonFile -Path $TransactionStatePath
    $temporaryCommitter = Join-Path $env:TEMP ('CityPC_Updater_commit_' + [string]$state.transactionId + '.ps1')
    Copy-Item -LiteralPath $PSCommandPath -Destination $temporaryCommitter -Force
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $temporaryCommitter + '"'),
        '-Mode', 'Commit', '-StatePath', ('"' + $TransactionStatePath + '"'), '-ParentPid', [string]$CallerProcessId
    ) -join ' '
    Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden | Out-Null
}

function Resolve-StaleTransaction {
    param([string]$StateRoot)
    $transactionRoot = Join-Path $StateRoot 'transactions'
    if (-not (Test-Path -LiteralPath $transactionRoot -PathType Container)) { return $false }
    $stateFiles = @(Get-ChildItem -LiteralPath $transactionRoot -Filter 'state.json' -File -Recurse | Sort-Object LastWriteTimeUtc -Descending)
    $active = @()
    foreach ($file in $stateFiles) {
        try {
            $state = Read-JsonFile -Path $file.FullName
            $status = [string]$state.status
            if ($status -ceq 'staged' -or $status -ceq 'committing' -or $status -ceq 'committed') {
                $active += [ordered]@{ file = $file.FullName; state = $state }
            }
        } catch {
            continue
        }
    }
    if ($active.Count -eq 0) { return $false }
    if ($active.Count -gt 1) {
        foreach ($item in $active) {
            Set-StateStatus -Path ([string]$item.file) -Status 'quarantined' `
                -Reason 'Multiples transacciones activas; no se eligio una automaticamente.' | Out-Null
        }
        throw 'Multiples transacciones activas fueron puestas en cuarentena.'
    }

    $selected = $active[0]
    $selectedState = $selected.state
    if ([string]$selectedState.status -ceq 'committed') {
        $valid = $true
        foreach ($entry in @($selectedState.entries)) {
            if (-not (Test-Path -LiteralPath ([string]$entry.targetPath) -PathType Leaf) -or
                (Get-Sha256 -Path ([string]$entry.targetPath)) -cne [string]$entry.expectedSha256) {
                $valid = $false
            }
        }
        if ($valid) {
            Set-StateStatus -Path ([string]$selected.file) -Status 'succeeded' `
                -Reason 'Recuperado tras corte antes de confirmar sentinel.' | Out-Null
            return $false
        }
        $selectedState.status = 'committing'
        Write-JsonAtomic -Path ([string]$selected.file) -Value $selectedState
    }
    Start-Committer -TransactionStatePath ([string]$selected.file) -CallerProcessId (Get-CallerProcessId)
    return $true
}

function Restore-Bundle {
    param($State)
    foreach ($entry in @($State.entries)) {
        $target = [string]$entry.targetPath
        $previous = [string]$entry.previousPath
        if ([bool]$entry.existedBefore) {
            $oldHash = ([string]$entry.oldSha256).ToLowerInvariant()
            $durable = [string]$entry.durableBackupPath
            $restoreSource = ''
            if (Test-FileSha256 -Path $previous -ExpectedHash $oldHash) {
                $restoreSource = $previous
            } elseif (Test-FileSha256 -Path $durable -ExpectedHash $oldHash) {
                # PREVIOUS_CORRUPT_DURABLE_FALLBACK: never trust a present but corrupt previous copy.
                $restoreSource = $durable
            } else {
                throw ('Rollback sin copia verificada: ' + [string]$entry.file)
            }

            $restoreTemporary = $target + '.citypc-restore-' + [string]$State.transactionId
            Remove-Item -LiteralPath $restoreTemporary -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath $restoreSource -Destination $restoreTemporary -Force
            if (-not (Test-FileSha256 -Path $restoreTemporary -ExpectedHash $oldHash)) {
                Remove-Item -LiteralPath $restoreTemporary -Force -ErrorAction SilentlyContinue
                throw ('Rollback temporal no verificable: ' + [string]$entry.file)
            }
            Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
            Move-Item -LiteralPath $restoreTemporary -Destination $target -Force
        } else {
            Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
        }
    }
    foreach ($entry in @($State.entries)) {
        $target = [string]$entry.targetPath
        if ([bool]$entry.existedBefore) {
            if (-not (Test-Path -LiteralPath $target -PathType Leaf) -or
                (Get-Sha256 -Path $target) -cne [string]$entry.oldSha256) {
                throw ('Rollback no verificable: ' + [string]$entry.file)
            }
        } elseif (Test-Path -LiteralPath $target -PathType Leaf) {
            throw ('Rollback dejo un archivo que antes no existia: ' + [string]$entry.file)
        }
    }
}

function Start-CallerRecovery {
    param($State)
    $callerEntries = @($State.entries | Where-Object { [string]$_.id -ceq [string]$State.callerToolId })
    if ($callerEntries.Count -ne 1 -or -not (Test-Path -LiteralPath ([string]$State.callerPath) -PathType Leaf)) {
        throw 'Caller de recovery ausente o ambiguo.'
    }
    $callerHash = Get-Sha256 -Path ([string]$State.callerPath)
    if ($callerHash -cne [string]$callerEntries[0].oldSha256 -and
        $callerHash -cne [string]$callerEntries[0].expectedSha256) {
        throw 'Caller de recovery no coincide con copia conocida.'
    }
    $transactionDir = Split-Path -Parent $StatePath
    $recoveryCmd = Join-Path $transactionDir 'recovery.cmd'
    $recoveryText = "@echo off`r`ncall `"$([string]$State.callerPath)`" --citypc-update-recovery `"$([string]$State.transactionId)`"`r`nexit /b %errorlevel%`r`n"
    [IO.File]::WriteAllText($recoveryCmd, $recoveryText, [Text.Encoding]::ASCII)
    $processInfo = New-Object Diagnostics.ProcessStartInfo
    $processInfo.FileName = $env:ComSpec
    $processInfo.Arguments = '/d /c ""' + $recoveryCmd + '""'
    $processInfo.WorkingDirectory = [string]$State.root
    $processInfo.UseShellExecute = $false
    [Diagnostics.Process]::Start($processInfo) | Out-Null
}

function Confirm-Resume {
    param([string]$Root, [string]$Token, [string]$CallerToolId, [int]$CallerVersion, [string]$CallerPath)
    if (-not (Test-SafeToken -Value $Token)) { throw 'Sentinel de reinicio invalido.' }
    $transactionDir = Join-Path (Join-Path (Join-Path $Root '.citypc-updater') 'transactions') $Token
    $stateFile = Join-Path $transactionDir 'state.json'
    if (-not (Test-Path -LiteralPath $stateFile -PathType Leaf)) { throw 'Sentinel sin transaccion.' }
    $state = Read-JsonFile -Path $stateFile
    if ([string]$state.transactionId -cne $Token -or [string]$state.status -cne 'committed') {
        throw 'Transaccion no esta lista para confirmar.'
    }
    if ([string]$state.callerToolId -cne $CallerToolId -or [int]$state.restartCount -ne 1) {
        throw 'Sentinel no pertenece al proceso reiniciado.'
    }
    if ([IO.Path]::GetFullPath([string]$state.callerPath) -cne [IO.Path]::GetFullPath($CallerPath)) {
        throw 'Ruta reiniciada no coincide.'
    }
    foreach ($entry in @($state.entries)) {
        if (-not (Test-Path -LiteralPath ([string]$entry.targetPath) -PathType Leaf)) { throw 'Bundle reiniciado incompleto.' }
        if ((Get-Sha256 -Path ([string]$entry.targetPath)) -cne [string]$entry.expectedSha256) {
            throw ('SHA final invalido: ' + [string]$entry.file)
        }
        if ([string]$entry.kind -ceq 'bat' -and (Get-BatchVersion -Path ([string]$entry.targetPath)) -ne [int]$entry.expectedVersion) {
            throw ('Version final invalida: ' + [string]$entry.file)
        }
    }
    $callerEntry = @($state.entries | Where-Object { [string]$_.id -ceq $CallerToolId })
    if ($callerEntry.Count -ne 1 -or [int]$callerEntry[0].expectedVersion -ne $CallerVersion) {
        throw 'Version del caller no coincide con bundle.'
    }
    $state.status = 'succeeded'
    $state.updatedAtUtc = [DateTime]::UtcNow.ToString('o')
    $state.completedAtUtc = [DateTime]::UtcNow.ToString('o')
    Write-JsonAtomic -Path $stateFile -Value $state
    Set-Content -LiteralPath (Join-Path $transactionDir 'resume.ok') -Value $Token -Encoding ASCII
    Remove-Item -LiteralPath (Join-Path (Join-Path $Root '.citypc-updater') 'failure.json') -Force -ErrorAction SilentlyContinue
    Write-Status 'OK' ('Bundle V{0} confirmado; reinicio unico completado.' -f [int]$state.bundleVersion)
}

function Invoke-CheckMode {
    if (-not $script:AllowedTools.ContainsKey($ToolId)) { throw 'ToolId no permitido.' }
    if ($CurrentVersion -le 0) { throw 'CurrentVersion invalida.' }
    $fullTarget = [IO.Path]::GetFullPath($TargetPath)
    if (-not (Test-Path -LiteralPath $fullTarget -PathType Leaf)) { throw 'BAT caller ausente.' }
    if ([IO.Path]::GetFileName($fullTarget) -cne $script:AllowedTools[$ToolId]) { throw 'TargetPath no permitido.' }
    $root = Split-Path -Parent $fullTarget
    $stateRoot = Join-Path $root '.citypc-updater'

    if ($ResumeToken) {
        Confirm-Resume -Root $root -Token $ResumeToken -CallerToolId $ToolId `
            -CallerVersion $CurrentVersion -CallerPath $fullTarget
        return $script:ExitContinue
    }

    $lock = $null
    try {
        $lock = Open-UpdateLock -StateRoot $stateRoot
        if (Resolve-StaleTransaction -StateRoot $stateRoot) {
            Write-Status 'INFO' 'Recuperando una transaccion interrumpida; cierre seguro solicitado.'
            return $script:ExitCallerMustClose
        }
        $script:DownloadDeadlineUtc = [DateTime]::UtcNow.AddSeconds(90)
        $scratch = Join-Path (Join-Path $stateRoot 'checks') ([Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $scratch -Force | Out-Null
        $remote = Get-RemoteBundle -ScratchDirectory $scratch
        if ($null -eq $remote) {
            Write-Status 'AVISO' ('No se pudo verificar el channel. Se usa V{0}; ningun archivo fue cambiado.' -f $CurrentVersion)
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
            return $script:ExitContinue
        }

        $manifest = $remote.manifest
        $needsUpdate = $false
        foreach ($tool in @($manifest.tools)) {
            $localPath = Join-Path $root ([string]$tool.file)
            if (-not (Test-Path -LiteralPath $localPath -PathType Leaf)) { $needsUpdate = $true; continue }
            try {
                $localVersion = Get-BatchVersion -Path $localPath
            } catch {
                if ([string]$tool.id -ceq 'onedrive' -and
                    (Get-Sha256 -Path $localPath) -ceq ([string]$tool.legacySha256).ToLowerInvariant()) {
                    $localVersion = 1
                } else {
                    $needsUpdate = $true
                    continue
                }
            }
            if ($localVersion -gt [int]$tool.version) {
                throw ('Se encontro una version local mas nueva en ' + [string]$tool.file + '; no se permite downgrade.')
            }
            if ($localVersion -lt [int]$tool.version -or
                (Get-Sha256 -Path $localPath) -cne ([string]$tool.sha256).ToLowerInvariant()) {
                $needsUpdate = $true
            }
        }
        $localUpdater = Join-Path $root 'CityPC_Updater.ps1'
        if (-not (Test-Path -LiteralPath $localUpdater -PathType Leaf) -or
            (Get-Sha256 -Path $localUpdater) -cne ([string]$manifest.updater.sha256).ToLowerInvariant()) {
            $needsUpdate = $true
        }
        if (-not $needsUpdate) {
            Write-Status 'OK' ('Bundle V{0} completo y verificado.' -f [int]$manifest.bundleVersion)
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
            return $script:ExitContinue
        }

        $failurePath = Join-Path $stateRoot 'failure.json'
        if (Get-FailureBlock -Path $failurePath -BundleVersion ([int]$manifest.bundleVersion) `
            -ManifestHash ([string]$remote.manifestSha256) {
            Write-Status 'AVISO' ('Bundle V{0} bloqueado 15 minutos tras rollback; se usa la copia actual.' -f [int]$manifest.bundleVersion)
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
            return $script:ExitContinue
        }

        $transactionId = [Guid]::NewGuid().ToString('N')
        $transactionDir = Join-Path (Join-Path $stateRoot 'transactions') $transactionId
        $stagedDir = Join-Path $transactionDir 'staged'
        $previousDir = Join-Path $transactionDir 'previous'
        $backupDir = Join-Path (Join-Path (Join-Path $stateRoot 'backups') ('bundle-v' + [int]$manifest.bundleVersion)) $transactionId
        New-Item -ItemType Directory -Path $stagedDir, $previousDir, $backupDir -Force | Out-Null

        $entries = @()
        foreach ($tool in @($manifest.tools)) {
            $staged = Join-Path $stagedDir ([string]$tool.file)
            if (-not (Invoke-Download -Url ([string]$remote.tagBase + '/' + [string]$tool.file) -Destination $staged)) {
                throw ('No se pudo descargar ' + [string]$tool.file)
            }
            Assert-DownloadedFile -Path $staged -Hash ([string]$tool.sha256) `
                -MinBytes ([int64]$tool.minBytes) -MaxBytes ([int64]$tool.maxBytes)
            Assert-BatchArtifact -Path $staged -Version ([int]$tool.version)
            $target = Join-Path $root ([string]$tool.file)
            $exists = Test-Path -LiteralPath $target -PathType Leaf
            $backup = Join-Path $backupDir ([string]$tool.file)
            if ($exists) {
                $oldHash = Get-Sha256 -Path $target
                Copy-Item -LiteralPath $target -Destination $backup
                if ((Get-Sha256 -Path $backup) -cne $oldHash) { throw ('Backup no verificable: ' + [string]$tool.file) }
            } else {
                $oldHash = ''
            }
            $entries += [ordered]@{
                id = [string]$tool.id; kind = 'bat'; file = [string]$tool.file
                targetPath = $target; stagedPath = $staged; previousPath = (Join-Path $previousDir ([string]$tool.file))
                durableBackupPath = $backup; existedBefore = [bool]$exists
                oldSha256 = $oldHash
                expectedSha256 = ([string]$tool.sha256).ToLowerInvariant(); expectedVersion = [int]$tool.version
            }
        }

        $stagedUpdater = Join-Path $stagedDir 'CityPC_Updater.ps1'
        if (-not (Invoke-Download -Url ([string]$remote.tagBase + '/CityPC_Updater.ps1') -Destination $stagedUpdater)) {
            throw 'No se pudo descargar CityPC_Updater.ps1.'
        }
        Assert-DownloadedFile -Path $stagedUpdater -Hash ([string]$manifest.updater.sha256) `
            -MinBytes ([int64]$manifest.updater.minBytes) -MaxBytes ([int64]$manifest.updater.maxBytes)
        $updaterExists = Test-Path -LiteralPath $localUpdater -PathType Leaf
        $updaterBackup = Join-Path $backupDir 'CityPC_Updater.ps1'
        if ($updaterExists) {
            $oldUpdaterHash = Get-Sha256 -Path $localUpdater
            Copy-Item -LiteralPath $localUpdater -Destination $updaterBackup
            if ((Get-Sha256 -Path $updaterBackup) -cne $oldUpdaterHash) { throw 'Backup updater no verificable.' }
        } else {
            $oldUpdaterHash = ''
        }
        $entries += [ordered]@{
            id = 'updater'; kind = 'updater'; file = 'CityPC_Updater.ps1'
            targetPath = $localUpdater; stagedPath = $stagedUpdater; previousPath = (Join-Path $previousDir 'CityPC_Updater.ps1')
            durableBackupPath = $updaterBackup; existedBefore = [bool]$updaterExists
            oldSha256 = $oldUpdaterHash
            expectedSha256 = ([string]$manifest.updater.sha256).ToLowerInvariant(); expectedVersion = [int]$manifest.updater.version
        }

        $callerRemote = @($manifest.tools | Where-Object { [string]$_.id -ceq $ToolId })
        if ($callerRemote.Count -ne 1) { throw 'Caller ausente del bundle.' }
        $stateFile = Join-Path $transactionDir 'state.json'
        $state = [ordered]@{
            schema = 2; transactionId = $transactionId; status = 'staged'
            bundleVersion = [int]$manifest.bundleVersion; manifestSha256 = [string]$remote.manifestSha256
            callerToolId = $ToolId; callerPath = $fullTarget; callerExpectedVersion = [int]$callerRemote[0].version
            root = $root; entries = $entries; restartCount = 0
            createdAtUtc = [DateTime]::UtcNow.ToString('o'); updatedAtUtc = [DateTime]::UtcNow.ToString('o'); reason = ''
        }
        Write-JsonAtomic -Path $stateFile -Value $state

        Start-Committer -TransactionStatePath $stateFile -CallerProcessId (Get-CallerProcessId)
        Write-Status 'OK' ('Bundle V{0}: cuatro archivos verificados. Cerrando para commit.' -f [int]$manifest.bundleVersion)
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
        return $script:ExitCallerMustClose
    } finally {
        if ($null -ne $lock) { $lock.Dispose() }
    }
}

function Invoke-CommitMode {
    if (-not (Test-Path -LiteralPath $StatePath -PathType Leaf)) { throw 'Estado de transaccion ausente.' }
    $state = Read-JsonFile -Path $StatePath
    if (-not (Test-SafeToken -Value ([string]$state.transactionId))) { throw 'ID de transaccion invalido.' }
    if ([string]$state.status -cne 'staged' -and [string]$state.status -cne 'committing') {
        throw 'Transaccion no esta recuperable.'
    }
    if (-not (Wait-ForProcessExit -ProcessId $ParentPid -TimeoutSeconds 45)) {
        $state = Set-StateStatus -Path $StatePath -Status 'failed' -Reason 'El BAT anterior no cerro en 45 segundos.'
        Set-FailureBlock -Path (Join-Path (Join-Path ([string]$state.root) '.citypc-updater') 'failure.json') `
            -BundleVersion ([int]$state.bundleVersion) -ManifestHash ([string]$state.manifestSha256) `
            -Reason 'El BAT anterior no cerro.'
        # PARENT_TIMEOUT_NO_RELAUNCH: el caller sigue vivo; relanzarlo duplicaria la preparacion.
        return $script:ExitContinue
    }

    $stateRoot = Join-Path ([string]$state.root) '.citypc-updater'
    $lock = $null
    try {
        $lock = Open-UpdateLock -StateRoot $stateRoot
        $state = Read-JsonFile -Path $StatePath
        if ([string]$state.status -ceq 'committing') {
            Restore-Bundle -State $state
            $state = Set-StateStatus -Path $StatePath -Status 'rolled_back' -Reason 'Rollback tras corte durante commit.'
            Set-FailureBlock -Path (Join-Path $stateRoot 'failure.json') -BundleVersion ([int]$state.bundleVersion) `
                -ManifestHash ([string]$state.manifestSha256) -Reason 'Corte durante commit.'
            Start-CallerRecovery -State $state
            return $script:ExitContinue
        }
        if ([string]$state.status -cne 'staged') { throw 'Estado cambio antes del commit.' }
        foreach ($entry in @($state.entries)) {
            $exists = Test-Path -LiteralPath ([string]$entry.targetPath) -PathType Leaf
            if ([bool]$entry.existedBefore -ne [bool]$exists) { throw ('Existencia local cambio: ' + [string]$entry.file) }
            if ($exists -and (Get-Sha256 -Path ([string]$entry.targetPath)) -cne [string]$entry.oldSha256) {
                throw ('Archivo local cambio: ' + [string]$entry.file)
            }
        }
        Set-StateStatus -Path $StatePath -Status 'committing' | Out-Null
        $state = Read-JsonFile -Path $StatePath

        try {
            foreach ($entry in @($state.entries)) {
                if ([bool]$entry.existedBefore) {
                    Move-Item -LiteralPath ([string]$entry.targetPath) -Destination ([string]$entry.previousPath)
                }
                Move-Item -LiteralPath ([string]$entry.stagedPath) -Destination ([string]$entry.targetPath)
            }
            foreach ($entry in @($state.entries)) {
                if ((Get-Sha256 -Path ([string]$entry.targetPath)) -cne [string]$entry.expectedSha256) {
                    throw ('Verificacion final fallo: ' + [string]$entry.file)
                }
                if ([string]$entry.kind -ceq 'bat' -and
                    (Get-BatchVersion -Path ([string]$entry.targetPath)) -ne [int]$entry.expectedVersion) {
                    throw ('Version final fallo: ' + [string]$entry.file)
                }
            }
        } catch {
            Restore-Bundle -State $state
            Set-StateStatus -Path $StatePath -Status 'rolled_back' -Reason $_.Exception.Message | Out-Null
            Set-FailureBlock -Path (Join-Path $stateRoot 'failure.json') -BundleVersion ([int]$state.bundleVersion) `
                -ManifestHash ([string]$state.manifestSha256) -Reason $_.Exception.Message
            Start-CallerRecovery -State $state
            return $script:ExitContinue
        }

        $state = Read-JsonFile -Path $StatePath
        $state.status = 'committed'; $state.restartCount = 1
        $state.updatedAtUtc = [DateTime]::UtcNow.ToString('o'); $state.committedAtUtc = [DateTime]::UtcNow.ToString('o')
        Write-JsonAtomic -Path $StatePath -Value $state

        $transactionDir = Split-Path -Parent $StatePath
        $restartCmd = Join-Path $transactionDir 'restart.cmd'
        $restartText = "@echo off`r`ncall `"$([string]$state.callerPath)`" --citypc-update-resume `"$([string]$state.transactionId)`"`r`nexit /b %errorlevel%`r`n"
        [IO.File]::WriteAllText($restartCmd, $restartText, [Text.Encoding]::ASCII)
        $processInfo = New-Object Diagnostics.ProcessStartInfo
        $processInfo.FileName = $env:ComSpec
        $processInfo.Arguments = '/d /c ""' + $restartCmd + '""'
        $processInfo.WorkingDirectory = [string]$state.root
        $processInfo.UseShellExecute = $false
        $restarted = [Diagnostics.Process]::Start($processInfo)

        $resumeOk = Join-Path $transactionDir 'resume.ok'
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        while ([DateTime]::UtcNow -lt $deadline) {
            if (Test-Path -LiteralPath $resumeOk -PathType Leaf) { return $script:ExitContinue }
            Start-Sleep -Milliseconds 500
        }

        if ($null -ne $restarted -and -not $restarted.HasExited) {
            Stop-Process -Id $restarted.Id -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }
        $state = Read-JsonFile -Path $StatePath
        Restore-Bundle -State $state
        Set-StateStatus -Path $StatePath -Status 'rolled_back' -Reason 'Sentinel no confirmado en 60 segundos.' | Out-Null
        Set-FailureBlock -Path (Join-Path $stateRoot 'failure.json') -BundleVersion ([int]$state.bundleVersion) `
            -ManifestHash ([string]$state.manifestSha256) -Reason 'Sentinel no confirmado.'
        Start-CallerRecovery -State $state
        return $script:ExitContinue
    } finally {
        if ($null -ne $lock) { $lock.Dispose() }
        $committerName = [IO.Path]::GetFileName($PSCommandPath)
        $committerDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($PSCommandPath))
        $expectedTempDirectory = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')
        if ($committerName -like 'CityPC_Updater_commit_*.ps1' -and
            $committerDirectory -ieq $expectedTempDirectory) {
            Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$result = $script:ExitContinue
try {
    if ($Mode -eq 'Commit') { $result = Invoke-CommitMode } else { $result = Invoke-CheckMode }
} catch {
    if ($Mode -eq 'Commit' -and $StatePath -and (Test-Path -LiteralPath $StatePath -PathType Leaf)) {
        try {
            $failedState = Read-JsonFile -Path $StatePath
            if ([string]$failedState.status -ceq 'committing') {
                Restore-Bundle -State $failedState
            }
            $failedState = Set-StateStatus -Path $StatePath -Status 'failed' -Reason $_.Exception.Message
            $failedRoot = Join-Path ([string]$failedState.root) '.citypc-updater'
            Set-FailureBlock -Path (Join-Path $failedRoot 'failure.json') `
                -BundleVersion ([int]$failedState.bundleVersion) `
                -ManifestHash ([string]$failedState.manifestSha256) -Reason $_.Exception.Message
            Start-CallerRecovery -State $failedState
        } catch {
            # El committer no oculta el error original; no intenta reemplazos adicionales.
        }
    }
    Write-Status 'AVISO' ('Actualizacion omitida de forma segura: ' + $_.Exception.Message)
    if ($Mode -eq 'Check' -and $ResumeToken) {
        Write-Status 'ERROR' 'El reinicio no fue confirmado; se detiene la herramienta para permitir rollback.'
        $result = $script:ExitResumeFailed
    } elseif ($Mode -eq 'Check' -and $_.Exception.Message -match 'Multiples transacciones activas') {
        Write-Status 'ERROR' 'Estado ambiguo en cuarentena; no se ejecutan tareas hasta revisarlo.'
        $result = $script:ExitUnsafeState
    } else {
        $result = $script:ExitContinue
    }
}
exit $result
