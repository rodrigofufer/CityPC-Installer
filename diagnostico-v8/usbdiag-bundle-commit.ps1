[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Commit', 'Confirm', 'Recover')]
    [string]$Action,
    [Parameter(Mandatory = $true)][string]$ScriptDir,
    [Parameter(Mandatory = $true)][string]$TransactionDir,
    [string]$LockDir,
    [int]$ParentPid = 0,
    [string]$Token,
    [long]$RemoteVersion = 0,
    [string]$BatName,
    [string]$HelperName,
    [string]$CommitterName,
    [string]$DownloadedBat,
    [string]$DownloadedHelper,
    [string]$DownloadedCommitter,
    [string]$ManifestPath,
    [string]$ManifestExpectedSha256,
    [string]$ChannelPath,
    [string]$ValidatedVersionPath,
    [string]$CommitPath,
    [string]$ManifestExpectedHashPath
)

$ErrorActionPreference = 'Stop'
$stateSchema = 'citypc.usbdiag.update-state.v1'
$manifestSchema = 'citypc.usbdiag.update-manifest.v1'
$statePath = Join-Path $TransactionDir 'state.json'
$newDir = Join-Path $TransactionDir 'new'
$backupDir = Join-Path $TransactionDir 'backup'
$privateNames = @('usbdiag.shared.local.cmd', 'wifi.local.cmd')
$mutated = $false
$rollbackConfirmed = $false
$parentExited = $false

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Assert-SafeName([string]$Name) {
    if ([IO.Path]::GetFileName($Name) -cne $Name -or [string]::IsNullOrWhiteSpace($Name)) {
        throw 'unsafe bundle filename'
    }
    if ($privateNames -contains $Name) {
        throw 'private configuration is not updateable'
    }
}

function Write-State([System.Collections.IDictionary]$State) {
    $temporary = Join-Path $TransactionDir 'state.next.json'
    $State | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $statePath -Force
}

function Read-State {
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        throw 'transaction state is missing'
    }
    $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$state.schema -cne $stateSchema) { throw 'transaction state schema mismatch' }
    if ([string]$state.token -notmatch '^[0-9a-f]{32}$') { throw 'transaction token mismatch' }
    return $state
}

function Get-StateFiles([pscustomobject]$State) {
    $properties = @($State.files.PSObject.Properties)
    if ($properties.Count -ne 3) { throw 'transaction file count mismatch' }
    foreach ($property in $properties) { Assert-SafeName ([string]$property.Name) }
    return $properties
}

function Wait-ForParentExit {
    if ($ParentPid -le 0) { throw 'parent pid is invalid' }
    $deadline = [DateTime]::UtcNow.AddSeconds(60)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($null -eq (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) {
            $script:parentExited = $true
            return
        }
        Start-Sleep -Milliseconds 200
    }
    throw 'parent BAT did not exit within 60 seconds'
}

function Write-RecoverySentinel([string]$RecoveryToken) {
    $path = Join-Path ([IO.Path]::GetTempPath()) ("citypc_diag_recovered_{0}.ok" -f $RecoveryToken)
    [IO.File]::WriteAllText($path, $RecoveryToken, [Text.Encoding]::ASCII)
}

function Set-LockOwner {
    if ([string]::IsNullOrWhiteSpace($LockDir) -or -not (Test-Path -LiteralPath $LockDir -PathType Container)) {
        throw 'update lock is missing'
    }
    $ownerPath = Join-Path $LockDir 'owner.pid'
    [IO.File]::WriteAllText($ownerPath, ([string]$PID), [Text.Encoding]::ASCII)
}

function Start-Bat([string]$Path, [string[]]$Arguments) {
    $joined = '""' + $Path + '" ' + ($Arguments -join ' ') + '"'
    Start-Process -FilePath $env:ComSpec -ArgumentList @('/d', '/c', $joined)
}

function Remove-Lock {
    if (-not [string]::IsNullOrWhiteSpace($LockDir) -and (Test-Path -LiteralPath $LockDir)) {
        Remove-Item -LiteralPath $LockDir -Recurse -Force
    }
}

function Restore-FromState([pscustomobject]$State) {
    $errors = @()
    foreach ($property in (Get-StateFiles $State)) {
        $name = [string]$property.Name
        $backup = Join-Path $backupDir $name
        $installed = Join-Path $ScriptDir $name
        if ($property.Value.backupExisted -isnot [bool]) {
            $errors += $name
            continue
        }
        if ([bool]$property.Value.backupExisted) {
            try { Copy-Item -LiteralPath $backup -Destination $installed -Force } catch { $errors += $name }
        } else {
            try { if (Test-Path -LiteralPath $installed) { Remove-Item -LiteralPath $installed -Force } } catch { $errors += $name }
        }
    }
    foreach ($property in (Get-StateFiles $State)) {
        $name = [string]$property.Name
        try {
            $installed = Join-Path $ScriptDir $name
            if ([bool]$property.Value.backupExisted) {
                $expected = ([string]$property.Value.backupSha256).ToLowerInvariant()
                if ($expected -notmatch '^[0-9a-f]{64}$' -or (Get-Sha256 $installed) -cne $expected) {
                    $errors += $name
                }
            } elseif (Test-Path -LiteralPath $installed) {
                $errors += $name
            }
        } catch { $errors += $name }
    }
    if ($errors.Count -gt 0) {
        throw ('rollback could not be confirmed: ' + (($errors | Select-Object -Unique) -join ', '))
    }
    $script:rollbackConfirmed = $true
}

function Invoke-Confirm {
    $state = Read-State
    if ([string]$state.token -cne $Token) { throw 'confirmation token mismatch' }
    if ([long]$state.version -ne $RemoteVersion) { throw 'confirmation version mismatch' }
    if ([string]$state.phase -cne 'installed-awaiting-confirmation') {
        throw 'transaction is not awaiting confirmation'
    }
    foreach ($property in (Get-StateFiles $state)) {
        $name = [string]$property.Name
        $expected = ([string]$property.Value.expectedSha256).ToLowerInvariant()
        if ($expected -notmatch '^[0-9a-f]{64}$' -or (Get-Sha256 (Join-Path $ScriptDir $name)) -cne $expected) {
            throw "confirmation hash mismatch for $name"
        }
    }
    $transactionManifest = Join-Path $TransactionDir 'manifest.json'
    if ((Get-Sha256 $transactionManifest) -cne ([string]$state.manifestSha256).ToLowerInvariant()) {
        throw 'confirmation manifest hash mismatch'
    }
    $batPath = Join-Path $ScriptDir ([string]$state.roles.bat)
    $batText = [IO.File]::ReadAllText($batPath)
    $match = [regex]::Match($batText, '(?mi)^\s*set\s+"LOCAL_VER=([0-9]+)"\s*$')
    if (-not $match.Success -or [long]$match.Groups[1].Value -ne $RemoteVersion) {
        throw 'confirmed BAT local version mismatch'
    }
    $confirmedState = [ordered]@{
        schema = [string]$state.schema
        token = [string]$state.token
        version = [long]$state.version
        manifestSha256 = [string]$state.manifestSha256
        phase = 'confirmed'
        roles = $state.roles
        files = $state.files
    }
    Write-State $confirmedState
    $completedDir = $TransactionDir + '.confirmed-' + $Token
    Move-Item -LiteralPath $TransactionDir -Destination $completedDir
    try { Remove-Item -LiteralPath $completedDir -Recurse -Force } catch {}
}

function Invoke-Recover {
    Wait-ForParentExit
    $state = Read-State
    $recoveryToken = [string]$state.token
    $batPath = Join-Path $ScriptDir ([string]$state.roles.bat)
    if ([string]$state.phase -eq 'confirmed') {
        foreach ($property in (Get-StateFiles $state)) {
            $name = [string]$property.Name
            $expected = ([string]$property.Value.expectedSha256).ToLowerInvariant()
            if ((Get-Sha256 (Join-Path $ScriptDir $name)) -cne $expected) {
                throw "confirmed recovery hash mismatch for $name"
            }
        }
        $transactionManifest = Join-Path $TransactionDir 'manifest.json'
        if ((Get-Sha256 $transactionManifest) -cne ([string]$state.manifestSha256).ToLowerInvariant()) {
            throw 'confirmed recovery manifest mismatch'
        }
    } else {
        Restore-FromState $state
    }
    Remove-Item -LiteralPath $TransactionDir -Recurse -Force
    Remove-Lock
    Write-RecoverySentinel $recoveryToken
    Start-Bat $batPath @('--recovered', $recoveryToken)
}

function Invoke-Commit {
    foreach ($required in @($LockDir, $Token, $BatName, $HelperName, $CommitterName, $DownloadedBat, $DownloadedHelper, $DownloadedCommitter, $ManifestPath, $ManifestExpectedSha256)) {
        if ([string]::IsNullOrWhiteSpace([string]$required)) { throw 'commit parameter is missing' }
    }
    if ($Token -notmatch '^[0-9a-f]{32}$') { throw 'commit token is invalid' }
    if ($RemoteVersion -le 0) { throw 'remote version is invalid' }
    if ($ManifestExpectedSha256 -notmatch '^[0-9a-f]{64}$' -or (Get-Sha256 $ManifestPath) -cne $ManifestExpectedSha256.ToLowerInvariant()) {
        throw 'anchored manifest hash mismatch'
    }
    if (Test-Path -LiteralPath $TransactionDir) { throw 'unfinished transaction already exists' }

    $names = @($BatName, $HelperName, $CommitterName)
    foreach ($name in $names) { Assert-SafeName $name }
    if (($names | Select-Object -Unique).Count -ne 3) { throw 'bundle names are not unique' }
    $downloads = @{
        $BatName = $DownloadedBat
        $HelperName = $DownloadedHelper
        $CommitterName = $DownloadedCommitter
    }

    $manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$manifest.schema -cne $manifestSchema) { throw 'manifest schema mismatch' }
    if ([long]$manifest.version -ne $RemoteVersion) { throw 'manifest version mismatch' }
    if (@($manifest.files.PSObject.Properties).Count -ne 3) { throw 'manifest file count mismatch' }

    foreach ($name in $names) {
        $entry = $manifest.files.PSObject.Properties[$name]
        if ($null -eq $entry) { throw "manifest missing $name" }
        $expected = ([string]$entry.Value).ToLowerInvariant()
        if ($expected -notmatch '^[0-9a-f]{64}$' -or (Get-Sha256 $downloads[$name]) -cne $expected) {
            throw "download hash mismatch for $name"
        }
    }

    New-Item -ItemType Directory -Path $newDir -Force | Out-Null
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Copy-Item -LiteralPath $ManifestPath -Destination (Join-Path $TransactionDir 'manifest.json')

    $stateFiles = [ordered]@{}
    foreach ($name in $names) {
        $installed = Join-Path $ScriptDir $name
        $newPath = Join-Path $newDir $name
        $backupPath = Join-Path $backupDir $name
        Copy-Item -LiteralPath $downloads[$name] -Destination $newPath
        $expected = ([string]$manifest.files.PSObject.Properties[$name].Value).ToLowerInvariant()
        $backupExisted = Test-Path -LiteralPath $installed -PathType Leaf
        $backupHash = $null
        if ($backupExisted) {
            Copy-Item -LiteralPath $installed -Destination $backupPath
            $backupHash = Get-Sha256 $backupPath
        }
        if ((Get-Sha256 $newPath) -cne $expected) {
            throw "transaction staging mismatch for $name"
        }
        if ($backupExisted -and (Get-Sha256 $installed) -cne $backupHash) {
            throw "transaction backup mismatch for $name"
        }
        $stateFiles[$name] = [ordered]@{
            expectedSha256 = $expected
            backupExisted = [bool]$backupExisted
            backupSha256 = $backupHash
        }
    }

    $state = [ordered]@{
        schema = $stateSchema
        token = $Token
        version = $RemoteVersion
        manifestSha256 = $ManifestExpectedSha256.ToLowerInvariant()
        phase = 'prepared'
        roles = [ordered]@{ bat = $BatName; helper = $HelperName; committer = $CommitterName }
        files = $stateFiles
    }
    Write-State $state
    Wait-ForParentExit

    $state.phase = 'installing'
    Write-State $state
    foreach ($name in @($HelperName, $CommitterName, $BatName)) {
        $copied = $false
        for ($attempt = 1; $attempt -le 20; $attempt++) {
            try {
                Copy-Item -LiteralPath (Join-Path $newDir $name) -Destination (Join-Path $ScriptDir $name) -Force
                $copied = $true
                break
            } catch { Start-Sleep -Milliseconds 250 }
        }
        if (-not $copied) { throw "copy failed for $name" }
        $script:mutated = $true
    }

    foreach ($name in $names) {
        $expected = ([string]$state.files[$name].expectedSha256).ToLowerInvariant()
        if ((Get-Sha256 (Join-Path $ScriptDir $name)) -cne $expected) {
            throw "installed hash mismatch for $name"
        }
    }

    $state.phase = 'installed-awaiting-confirmation'
    Write-State $state
    Remove-Lock
    Start-Bat (Join-Path $ScriptDir $BatName) @('--updated', $Token, ([string]$RemoteVersion))
}

try {
    if ($Action -eq 'Commit' -or $Action -eq 'Recover') { Set-LockOwner }
    switch ($Action) {
        'Commit' { Invoke-Commit }
        'Confirm' { Invoke-Confirm }
        'Recover' { Invoke-Recover }
    }
    exit 0
} catch {
    $failure = $_.Exception.Message
    if ($Action -eq 'Commit' -and (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        try {
            $stateForRollback = Read-State
            if ($mutated -or [string]$stateForRollback.phase -ne 'prepared') {
                Restore-FromState $stateForRollback
            } else {
                $rollbackConfirmed = $true
            }
        } catch {
            $failure = $failure + '; ' + $_.Exception.Message
        }
    }
    if ($Action -eq 'Commit' -and -not $mutated -and -not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        $rollbackConfirmed = $true
        try { if (Test-Path -LiteralPath $TransactionDir) { Remove-Item -LiteralPath $TransactionDir -Recurse -Force } } catch {}
    }
    if ($Action -ne 'Confirm' -and $parentExited) {
        try { Remove-Lock } catch {}
    }

    if ($Action -eq 'Commit' -and $rollbackConfirmed -and $parentExited) {
        $recoveryToken = $Token
        $batPath = Join-Path $ScriptDir $BatName
        try { Remove-Item -LiteralPath $TransactionDir -Recurse -Force } catch {}
        Write-RecoverySentinel $recoveryToken
        Start-Bat $batPath @('--recovered', $recoveryToken)
    } else {
        Write-Host ''
        if (($Action -eq 'Commit' -or $Action -eq 'Recover') -and -not $parentExited) {
            Write-Host '[ERROR] El BAT padre no salio; no se libero el lock ni se relanzo otra copia.' -ForegroundColor Yellow
        }
        if (($Action -eq 'Commit' -or $Action -eq 'Recover') -and -not $rollbackConfirmed) {
            Write-Host '[ERROR CRITICO] No se confirmo el rollback completo del bundle.' -ForegroundColor Red
            Write-Host 'No ejecute el diagnostico. Conserve la transaccion y recopie el paquete completo.' -ForegroundColor Red
        } else {
            Write-Host '[ERROR] La actualizacion no pudo confirmarse.' -ForegroundColor Yellow
        }
        Write-Host $failure
        Read-Host 'Presione Enter para cerrar' | Out-Null
    }
    exit 40
} finally {
    if ($Action -eq 'Commit') {
        foreach ($path in @($DownloadedBat, $DownloadedHelper, $ManifestPath, $ChannelPath, $ValidatedVersionPath, $CommitPath, $ManifestExpectedHashPath, $DownloadedCommitter)) {
            try { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } } catch {}
        }
    }
}
