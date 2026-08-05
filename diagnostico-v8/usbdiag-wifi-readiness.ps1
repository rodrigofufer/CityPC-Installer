[CmdletBinding()]
param(
    [ValidateSet('Connect', 'Cleanup')]
    [string]$Action = 'Connect',
    [int]$ReadyTimeoutSeconds = 45
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Exit-WithCode {
    param(
        [int]$Code,
        [string]$Message
    )

    if ($Message) {
        [Console]::Error.WriteLine($Message)
    }
    exit $Code
}

function Invoke-NetshBounded {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,
        [int]$TimeoutMilliseconds = 15000
    )

    $process = $null
    try {
        $process = Start-Process -FilePath "$env:SystemRoot\System32\netsh.exe" `
            -ArgumentList $Arguments `
            -PassThru `
            -WindowStyle Hidden
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {
            try { $process.Kill() } catch {}
            return 124
        }
        return [int]$process.ExitCode
    } catch {
        return 125
    } finally {
        if ($null -ne $process) {
            $process.Dispose()
        }
    }
}

function Remove-DiagnosticWifiProfile {
    $profileArgument = 'name="{0}"' -f $env:WIFI_PROFILE_NAME
    return Invoke-NetshBounded -Arguments @('wlan', 'delete', 'profile', $profileArgument)
}

function Write-WlanProfile {
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $true
    $settings.CloseOutput = $true
    $namespace = 'http://www.microsoft.com/networking/WLAN/profile/v1'

    $writer = [System.Xml.XmlWriter]::Create($env:WIFI_XML, $settings)
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('WLANProfile', $namespace)
        $writer.WriteElementString('name', $namespace, $env:WIFI_PROFILE_NAME)
        $writer.WriteStartElement('SSIDConfig', $namespace)
        $writer.WriteStartElement('SSID', $namespace)
        $writer.WriteElementString('name', $namespace, $env:WIFI_SSID)
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteElementString('connectionType', $namespace, 'ESS')
        $writer.WriteElementString('connectionMode', $namespace, 'manual')
        $writer.WriteStartElement('MSM', $namespace)
        $writer.WriteStartElement('security', $namespace)
        $writer.WriteStartElement('authEncryption', $namespace)
        $writer.WriteElementString('authentication', $namespace, 'WPA2PSK')
        $writer.WriteElementString('encryption', $namespace, 'AES')
        $writer.WriteElementString('useOneX', $namespace, 'false')
        $writer.WriteEndElement()
        $writer.WriteStartElement('sharedKey', $namespace)
        $writer.WriteElementString('keyType', $namespace, 'passPhrase')
        $writer.WriteElementString('protected', $namespace, 'false')
        $writer.WriteElementString('keyMaterial', $namespace, $env:WIFI_PASS)
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    } finally {
        $writer.Dispose()
    }
}

function Get-ExpectedWifiState {
    $state = [ordered]@{
        ExactSsid = $false
        HasIpv4 = $false
    }

    $networkInformation = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]
    foreach ($profile in @($networkInformation::GetConnectionProfiles())) {
        if (-not $profile.IsWlanConnectionProfile) {
            continue
        }

        $connectedSsid = [string]$profile.WlanConnectionProfileDetails.GetConnectedSsid()
        if ($connectedSsid -cne $env:WIFI_SSID) {
            continue
        }

        $state.ExactSsid = $true
        $adapterId = [Guid]$profile.NetworkAdapter.NetworkAdapterId
        $adapter = $null
        foreach ($candidate in @(Get-NetAdapter -ErrorAction SilentlyContinue)) {
            if ([Guid]$candidate.InterfaceGuid -eq $adapterId) {
                $adapter = $candidate
                break
            }
        }
        if ($null -eq $adapter) {
            continue
        }

        foreach ($address in @(Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
            $ip = [string]$address.IPAddress
            if ($ip -and $ip -notmatch '^(0\.|127\.|169\.254\.)' -and $address.AddressState -notin @('Tentative', 'Duplicate')) {
                $state.HasIpv4 = $true
                break
            }
        }
    }

    return [pscustomobject]$state
}

function Assert-DnsAndHttps {
    if ([string]::IsNullOrWhiteSpace($env:GITHUB_CHANNEL_RAW) -or [string]::IsNullOrWhiteSpace($env:CHANNEL_FILE)) {
        throw 'No se definio el canal HTTPS de actualizacion.'
    }
    $channelUri = [Uri]($env:GITHUB_CHANNEL_RAW.TrimEnd('/') + '/' + $env:CHANNEL_FILE)
    $webhookUri = [Uri]$env:WEBHOOK_URL
    if ($channelUri.Scheme -cne 'https') {
        throw 'El canal de actualizacion no usa HTTPS.'
    }
    $hosts = @($channelUri.DnsSafeHost, $webhookUri.DnsSafeHost) | Select-Object -Unique

    foreach ($hostName in $hosts) {
        $parsedAddress = $null
        if ([System.Net.IPAddress]::TryParse($hostName, [ref]$parsedAddress)) {
            continue
        }
        Resolve-DnsName -Name $hostName -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null
    }

    Assert-TlsHandshake -Uri $webhookUri -TimeoutMilliseconds 10000

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $channel = $null
    $lastFailure = $null
    $deadline = [DateTime]::UtcNow.AddSeconds(35)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $remaining = [int][Math]::Floor(($deadline - [DateTime]::UtcNow).TotalSeconds)
            if ($remaining -le 0) { throw 'timeout total del canal' }
            $timeout = [Math]::Min(12, $remaining)
            $response = Invoke-WebRequest -Uri $channelUri.AbsoluteUri -Method Get -TimeoutSec $timeout -UseBasicParsing -ErrorAction Stop
            $channel = ([string]$response.Content) | ConvertFrom-Json
            $version = [string]$channel.version
            $commit = ([string]$channel.commit).ToLowerInvariant()
            $manifestHash = ([string]$channel.manifestSha256).ToLowerInvariant()
            if ([string]$channel.schema -cne 'citypc.usbdiag.update-channel.v1' -or
                $version -notmatch '^[0-9]{1,9}$' -or
                $commit -notmatch '^[0-9a-f]{40}$' -or
                $manifestHash -notmatch '^[0-9a-f]{64}$') {
                throw 'contrato invalido del canal'
            }
            break
        } catch {
            $lastFailure = $_.Exception.Message
            $channel = $null
            if ($attempt -lt 3) { Start-Sleep -Milliseconds 700 }
        }
    }
    if ($null -eq $channel) {
        throw ('La comprobacion HTTPS del canal fallo: ' + $lastFailure)
    }
}

function Assert-TlsHandshake {
    param(
        [Parameter(Mandatory = $true)]
        [Uri]$Uri,
        [int]$TimeoutMilliseconds = 10000
    )

    if ($Uri.Scheme -cne 'https') {
        throw 'El endpoint configurado no usa HTTPS.'
    }

    $port = if ($Uri.IsDefaultPort) { 443 } else { $Uri.Port }
    $client = [System.Net.Sockets.TcpClient]::new()
    $ssl = $null
    try {
        $connectResult = $client.BeginConnect($Uri.DnsSafeHost, $port, $null, $null)
        if (-not $connectResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
            throw 'Timeout TCP al endpoint configurado.'
        }
        $client.EndConnect($connectResult)

        $ssl = [System.Net.Security.SslStream]::new($client.GetStream(), $false)
        $authResult = $ssl.BeginAuthenticateAsClient($Uri.DnsSafeHost, $null, $null)
        if (-not $authResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
            throw 'Timeout TLS al endpoint configurado.'
        }
        $ssl.EndAuthenticateAsClient($authResult)
        if (-not $ssl.IsAuthenticated -or -not $ssl.IsEncrypted) {
            throw 'El endpoint no completo un handshake TLS autenticado y cifrado.'
        }
    } finally {
        if ($null -ne $ssl) {
            $ssl.Dispose()
        }
        $client.Dispose()
    }
}

if ([string]::IsNullOrWhiteSpace($env:WIFI_SSID) -or
    [string]::IsNullOrWhiteSpace($env:WIFI_PASS) -or
    [string]::IsNullOrWhiteSpace($env:WIFI_PROFILE_NAME) -or
    $env:WIFI_SSID.Length -gt 32 -or
    $env:WIFI_PASS.Length -lt 8 -or
    $env:WIFI_PASS.Length -gt 63 -or
    $env:WIFI_SSID.IndexOfAny([char[]]@('"', "`r", "`n")) -ge 0) {
    Exit-WithCode 41 'Configuracion Wi-Fi invalida: revise SSID y password local.'
}

if ($Action -eq 'Cleanup') {
    $cleanupCode = Remove-DiagnosticWifiProfile
    if ($cleanupCode -ne 0) {
        Exit-WithCode 49 'No se pudo retirar el perfil Wi-Fi temporal del equipo.'
    }
    exit 0
}

if ([string]::IsNullOrWhiteSpace($env:WIFI_XML)) {
    Exit-WithCode 41 'No se definio la ruta temporal segura del perfil Wi-Fi.'
}

try {
    $service = Get-Service -Name WlanSvc -ErrorAction Stop
    if ($service.Status -ne 'Running') {
        Start-Service -Name WlanSvc -ErrorAction Stop
        $service.WaitForStatus('Running', [TimeSpan]::FromSeconds(10))
    }
} catch {
    Exit-WithCode 42 'El servicio Wi-Fi de Windows no esta disponible o no pudo iniciarse.'
}

try {
    if (Test-Path -LiteralPath $env:WIFI_XML) {
        Remove-Item -LiteralPath $env:WIFI_XML -Force -ErrorAction Stop
    }
    Write-WlanProfile
} catch {
    Exit-WithCode 43 'No se pudo crear el perfil Wi-Fi temporal.'
}

$ready = $false
try {
    [void](Remove-DiagnosticWifiProfile)
    $addCode = Invoke-NetshBounded -Arguments @('wlan', 'add', 'profile', ('filename="{0}"' -f $env:WIFI_XML), 'user=current')
    if ($addCode -ne 0) {
        Exit-WithCode 44 'Windows rechazo el perfil Wi-Fi temporal.'
    }
    $connectCode = Invoke-NetshBounded -Arguments @('wlan', 'connect', ('name="{0}"' -f $env:WIFI_PROFILE_NAME), ('ssid="{0}"' -f $env:WIFI_SSID))
    if ($connectCode -ne 0) {
        Exit-WithCode 45 'Windows no acepto la orden de conexion a la red configurada.'
    }

    $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(10, [Math]::Min($ReadyTimeoutSeconds, 60)))
    $sawExactSsid = $false
    $sawIpv4 = $false
    while ([DateTime]::UtcNow -lt $deadline) {
        try {
            $state = Get-ExpectedWifiState
            $sawExactSsid = $sawExactSsid -or $state.ExactSsid
            $sawIpv4 = $sawIpv4 -or $state.HasIpv4
            if ($state.ExactSsid -and $state.HasIpv4) {
                break
            }
        } catch {}
        Start-Sleep -Seconds 2
    }

    if (-not $sawExactSsid) {
        Exit-WithCode 46 'No se confirmo conexion al SSID exacto dentro del tiempo limite.'
    }
    if (-not $sawIpv4) {
        Exit-WithCode 47 'La red Wi-Fi no entrego una direccion IPv4 valida dentro del tiempo limite.'
    }

    try {
        Assert-DnsAndHttps
    } catch {
        Exit-WithCode 48 'Wi-Fi conectado, pero fallo DNS o HTTPS hacia los servicios requeridos.'
    }

    $ready = $true
    exit 0
} finally {
    if (Test-Path -LiteralPath $env:WIFI_XML) {
        Remove-Item -LiteralPath $env:WIFI_XML -Force -ErrorAction SilentlyContinue
    }
    if (-not $ready) {
        [void](Remove-DiagnosticWifiProfile)
    }
}
