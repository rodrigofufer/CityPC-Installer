@echo off
cls
color 0B
setlocal EnableExtensions DisableDelayedExpansion

:: =========================================================
:: 0. UNA SOLA EJECUCION POR EQUIPO
:: =========================================================
set "INSTANCE_LOCK=%temp%\citypc_usbdiag_v8.lock"
set "INSTANCE_LOCK_HELD="
call :ACQUIRE_INSTANCE_LOCK
if errorlevel 1 goto :ERROR_ALREADY_RUNNING

:: =========================================================
:: 1. CONFIGURACION LOCAL PRIVADA (NUNCA VERSIONAR)
:: =========================================================
set "SHARED_CONFIG=%~dp0usbdiag.shared.local.cmd"
set "WIFI_CONFIG=%~dp0wifi.local.cmd"
if not exist "%SHARED_CONFIG%" goto :ERROR_CONFIG
if not exist "%WIFI_CONFIG%" goto :ERROR_CONFIG
call "%SHARED_CONFIG%"
if errorlevel 1 goto :ERROR_CONFIG
call "%WIFI_CONFIG%"
if errorlevel 1 goto :ERROR_CONFIG
if not defined WEBHOOK_URL goto :ERROR_CONFIG
if not defined CITYPC_USB_TOKEN goto :ERROR_CONFIG
if not defined WIFI_SSID goto :ERROR_CONFIG
if not defined WIFI_PASS goto :ERROR_CONFIG
if /I not "%WEBHOOK_URL:~0,8%"=="https://" goto :ERROR_CONFIG

:: =========================================================
:: VERSION LOCAL Y AUTO-UPDATE
:: =========================================================
set "LOCAL_VER=8"
set "GITHUB_CHANNEL_RAW=https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main"
set "REMOTE_FILE=Diagnostico_CityPC.bat"
set "REMOTE_HELPER_FILE=usbdiag-wifi-readiness.ps1"
set "REMOTE_COMMITTER_FILE=usbdiag-bundle-commit.ps1"
set "SCRIPT_DIR=%~dp0"
set "CHANNEL_FILE=update-channel.json"
set "MANIFEST_FILE=update-manifest.json"
set "WIFI_HELPER=%~dp0usbdiag-wifi-readiness.ps1"
set "BUNDLE_COMMITTER=%~dp0usbdiag-bundle-commit.ps1"
set "UPDATE_NONCE=%RANDOM%%RANDOM%"
set "UPDATE_CHANNEL_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_channel.json"
set "UPDATE_VALID_VERSION_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_version_valid.txt"
set "UPDATE_COMMIT_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_commit.txt"
set "UPDATE_MANIFEST_EXPECTED_HASH_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_manifest_expected.sha256"
set "UPDATE_MANIFEST_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_manifest.json"
set "UPDATE_BAT_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_update.bat"
set "UPDATE_HELPER_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_wifi_helper.ps1"
set "UPDATE_COMMITTER_PATH=%temp%\citypc_diag_%UPDATE_NONCE%_bundle_commit.ps1"
set "UPDATE_TRANSACTION_DIR=%~dp0.usbdiag-update-transaction"

set "WIFI_XML=%temp%\citypc_diag_%UPDATE_NONCE%_wifi.xml"
set "WIFI_PROFILE_NAME=CityPC-Diagnostico-%UPDATE_NONCE%"
set "WIFI_READY="
set "WIFI_CLEANUP_FAILED="
set "UPDATE_POST_ACTION="

:: =========================================================
:: RECUPERACION/CONFIRMACION DURABLE DE AUTO-UPDATE
:: =========================================================
if /I "%~1"=="--updated" goto :CONFIRM_INSTALLED_BUNDLE
if /I "%~1"=="--recovered" goto :CONFIRM_RECOVERED_BUNDLE
if not "%~1"=="" goto :ERROR_UPDATE_HANDSHAKE
if exist "%UPDATE_TRANSACTION_DIR%\state.json" goto :LAUNCH_UPDATE_RECOVERY
if exist "%UPDATE_TRANSACTION_DIR%" goto :CLEAN_STATELESS_UPDATE_TRANSACTION
goto :UPDATE_STARTUP_READY

:CLEAN_STATELESS_UPDATE_TRANSACTION
:: El committer siempre escribe state.json antes de mutar el bundle. Sin estado
:: solo hay staging incompleto de un proceso muerto y se puede retirar.
rmdir /S /Q "%UPDATE_TRANSACTION_DIR%" >nul 2>&1
if exist "%UPDATE_TRANSACTION_DIR%" goto :ERROR_UPDATE_RECOVERY
echo [INFO] Se retiro una preparacion de update interrumpida antes del commit.
goto :UPDATE_STARTUP_READY

:CONFIRM_INSTALLED_BUNDLE
if "%~2"=="" goto :ERROR_UPDATE_HANDSHAKE
if "%~3"=="" goto :ERROR_UPDATE_HANDSHAKE
if not "%~4"=="" goto :ERROR_UPDATE_HANDSHAKE
set "UPDATE_CONFIRM_TOKEN=%~2"
set "UPDATE_CONFIRM_VERSION=%~3"
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:UPDATE_CONFIRM_TOKEN -cnotmatch '^[0-9a-f]{32}$' -or $env:UPDATE_CONFIRM_VERSION -cnotmatch '^[0-9]{1,9}$' -or [int64]$env:UPDATE_CONFIRM_VERSION -ne [int64]$env:LOCAL_VER){exit 1}else{exit 0}" >nul 2>&1
if errorlevel 1 goto :ERROR_UPDATE_HANDSHAKE
if not exist "%UPDATE_TRANSACTION_DIR%\state.json" goto :ERROR_UPDATE_HANDSHAKE
if not exist "%BUNDLE_COMMITTER%" goto :ERROR_UPDATE_COMMITTER
powershell -NoProfile -ExecutionPolicy Bypass -File "%BUNDLE_COMMITTER%" -Action Confirm -ScriptDir "%SCRIPT_DIR%" -TransactionDir "%UPDATE_TRANSACTION_DIR%" -Token "%UPDATE_CONFIRM_TOKEN%" -RemoteVersion "%UPDATE_CONFIRM_VERSION%"
if errorlevel 1 goto :ERROR_UPDATE_CONFIRMATION
set "UPDATE_POST_ACTION=confirmed"
goto :UPDATE_STARTUP_READY

:CONFIRM_RECOVERED_BUNDLE
if "%~2"=="" goto :ERROR_UPDATE_HANDSHAKE
if not "%~3"=="" goto :ERROR_UPDATE_HANDSHAKE
set "UPDATE_RECOVERY_TOKEN=%~2"
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:UPDATE_RECOVERY_TOKEN -cmatch '^[0-9a-f]{32}$'){exit 0}else{exit 1}" >nul 2>&1
if errorlevel 1 goto :ERROR_UPDATE_HANDSHAKE
if exist "%UPDATE_TRANSACTION_DIR%" goto :ERROR_UPDATE_HANDSHAKE
set "UPDATE_RECOVERY_SENTINEL=%temp%\citypc_diag_recovered_%UPDATE_RECOVERY_TOKEN%.ok"
if not exist "%UPDATE_RECOVERY_SENTINEL%" goto :ERROR_UPDATE_HANDSHAKE
set "UPDATE_RECOVERY_SENTINEL_VALUE="
set /p UPDATE_RECOVERY_SENTINEL_VALUE=<"%UPDATE_RECOVERY_SENTINEL%"
if not "%UPDATE_RECOVERY_SENTINEL_VALUE%"=="%UPDATE_RECOVERY_TOKEN%" goto :ERROR_UPDATE_HANDSHAKE
del /F /Q "%UPDATE_RECOVERY_SENTINEL%" >nul 2>&1
if not exist "%WIFI_HELPER%" goto :ERROR_WIFI_HELPER
if exist "%BUNDLE_COMMITTER%" goto :RECOVERED_BUNDLE_READY
goto :ERROR_UPDATE_COMMITTER
:RECOVERED_BUNDLE_READY
set "UPDATE_POST_ACTION=recovered"
goto :UPDATE_STARTUP_READY

:LAUNCH_UPDATE_RECOVERY
set "UPDATE_RECOVERY_COMMITTER=%UPDATE_TRANSACTION_DIR%\new\usbdiag-bundle-commit.ps1"
if not exist "%UPDATE_RECOVERY_COMMITTER%" goto :ERROR_UPDATE_RECOVERY
call :CAPTURE_PARENT_PID
if errorlevel 1 goto :ERROR_UPDATE_RECOVERY
start "" powershell -NoProfile -ExecutionPolicy Bypass -File "%UPDATE_RECOVERY_COMMITTER%" -Action Recover -ScriptDir "%SCRIPT_DIR%" -TransactionDir "%UPDATE_TRANSACTION_DIR%" -LockDir "%INSTANCE_LOCK%" -ParentPid "%UPDATE_PARENT_PID%"
if errorlevel 1 goto :ERROR_UPDATE_RECOVERY
set "INSTANCE_LOCK_HELD="
exit

:UPDATE_STARTUP_READY

:: =========================================================
:: CONEXION WI-FI DETERMINISTA (antes del auto-update)
:: =========================================================
if not exist "%WIFI_HELPER%" goto :ERROR_WIFI_HELPER
echo.
echo Preparando Wi-Fi y comprobando SSID, IPv4, DNS y HTTPS...
powershell -NoProfile -ExecutionPolicy Bypass -File "%WIFI_HELPER%" -Action Connect -ReadyTimeoutSeconds 45
set "WIFI_RC=%errorlevel%"
if not "%WIFI_RC%"=="0" goto :ERROR_WIFI
set "WIFI_READY=1"
echo [OK] Wi-Fi confirmado en la red esperada, con IPv4, DNS y HTTPS.
echo.

:: El updater usa expansion retardada; los secretos ya no se interpolan.
setlocal EnableDelayedExpansion

:: Un post-update solo salta tras confirmar token, version, estado y hashes.
if defined UPDATE_POST_ACTION (
    echo.
    echo [INFO] Estado de actualizacion !UPDATE_POST_ACTION!; se omite una segunda comprobacion.
    echo.
    goto :skip_update_diag
)

cls
echo.
echo ============================================================
echo      SISTEMA DE DIAGNOSTICO CITYPC (V%LOCAL_VER%)
echo ============================================================
echo.
echo Verificando si hay una version nueva...
echo.

:: El canal pequeno vive en main; ancla version, commit inmutable y hash del manifest.
call :CLEAN_UPDATE_FILES
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $deadline=[DateTime]::UtcNow.AddSeconds(35); for($attempt=1;$attempt -le 3;$attempt++){try{$remaining=[int][Math]::Floor(($deadline-[DateTime]::UtcNow).TotalSeconds); if($remaining -le 0){throw 'deadline'}; $timeout=[Math]::Min(12,$remaining); Invoke-WebRequest -Uri ($env:GITHUB_CHANNEL_RAW + '/' + $env:CHANNEL_FILE) -OutFile $env:UPDATE_CHANNEL_PATH -TimeoutSec $timeout -UseBasicParsing; exit 0}catch{if($attempt -lt 3){Start-Sleep -Milliseconds 700}}}; exit 1" >nul 2>&1
if errorlevel 1 goto :UPDATE_UNAVAILABLE

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; try { $c=Get-Content -LiteralPath $env:UPDATE_CHANNEL_PATH -Raw -Encoding UTF8 ^| ConvertFrom-Json; if([string]$c.schema -cne 'citypc.usbdiag.update-channel.v1'){exit 10}; $v=[string]$c.version; $commit=([string]$c.commit).ToLowerInvariant(); $manifestHash=([string]$c.manifestSha256).ToLowerInvariant(); if($v -notmatch '^[0-9]{1,9}$' -or $commit -notmatch '^[0-9a-f]{40}$' -or $manifestHash -notmatch '^[0-9a-f]{64}$'){exit 11}; [IO.File]::WriteAllText($env:UPDATE_VALID_VERSION_PATH,$v,[Text.Encoding]::ASCII); [IO.File]::WriteAllText($env:UPDATE_COMMIT_PATH,$commit,[Text.Encoding]::ASCII); [IO.File]::WriteAllText($env:UPDATE_MANIFEST_EXPECTED_HASH_PATH,$manifestHash,[Text.Encoding]::ASCII); if([int64]$v -lt [int64]$env:LOCAL_VER){exit 21}; if([int64]$v -eq [int64]$env:LOCAL_VER){exit 20}; exit 0 } catch { exit 12 }" >nul 2>&1
set "VERSION_CHECK=!errorlevel!"
if "!VERSION_CHECK!"=="21" (
    echo [OK] El canal apunta a una version menor; no se permite downgrade.
    echo.
    call :CLEAN_UPDATE_FILES
    goto :skip_update_diag
)
if not "!VERSION_CHECK!"=="0" if not "!VERSION_CHECK!"=="20" goto :UPDATE_INVALID

set "REMOTE_VER="
set "BUNDLE_COMMIT="
set "EXPECTED_MANIFEST_HASH="
set /p REMOTE_VER=<"!UPDATE_VALID_VERSION_PATH!"
set /p BUNDLE_COMMIT=<"!UPDATE_COMMIT_PATH!"
set /p EXPECTED_MANIFEST_HASH=<"!UPDATE_MANIFEST_EXPECTED_HASH_PATH!"
if not defined REMOTE_VER goto :UPDATE_INVALID
if not defined BUNDLE_COMMIT goto :UPDATE_INVALID
if not defined EXPECTED_MANIFEST_HASH goto :UPDATE_INVALID
set "BUNDLE_RAW=https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/!BUNDLE_COMMIT!"
set "BUNDLE_FALLBACK_RAW=https://github.com/rodrigofufer/CityPC-Installer/raw/!BUNDLE_COMMIT!"

echo Descargando bundle anclado al commit !BUNDLE_COMMIT:~0,12!...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $deadline=[DateTime]::UtcNow.AddSeconds(100); $bases=@($env:BUNDLE_RAW,$env:BUNDLE_FALLBACK_RAW); function Get-Anchored([string]$name,[string]$output,[int]$perRequest){for($attempt=1;$attempt -le 3;$attempt++){try{$remaining=[int][Math]::Floor(($deadline-[DateTime]::UtcNow).TotalSeconds); if($remaining -le 0){throw 'bundle deadline'}; $timeout=[Math]::Min($perRequest,$remaining); $base=$bases[($attempt-1) %% $bases.Count]; Invoke-WebRequest -Uri ($base + '/' + $name) -OutFile $output -TimeoutSec $timeout -UseBasicParsing; return}catch{if($attempt -lt 3){Start-Sleep -Milliseconds 700}}}; throw ('download failed: ' + $name)}; try { Get-Anchored $env:MANIFEST_FILE $env:UPDATE_MANIFEST_PATH 20; Get-Anchored $env:REMOTE_FILE $env:UPDATE_BAT_PATH 30; Get-Anchored $env:REMOTE_HELPER_FILE $env:UPDATE_HELPER_PATH 30; Get-Anchored $env:REMOTE_COMMITTER_FILE $env:UPDATE_COMMITTER_PATH 30; exit 0 } catch { exit 1 }" >nul 2>&1
if errorlevel 1 goto :UPDATE_UNAVAILABLE

:: Validar manifest anclado, TODO el bundle y contrato antes de mutar.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; try { $manifestActual=(Get-FileHash -LiteralPath $env:UPDATE_MANIFEST_PATH -Algorithm SHA256).Hash.ToLowerInvariant(); if($manifestActual -cne $env:EXPECTED_MANIFEST_HASH){exit 29}; $m=Get-Content -LiteralPath $env:UPDATE_MANIFEST_PATH -Raw -Encoding UTF8 ^| ConvertFrom-Json; if([string]$m.schema -cne 'citypc.usbdiag.update-manifest.v1'){exit 30}; if([int64]$m.version -ne [int64]$env:REMOTE_VER){exit 31}; $batEntry=$m.files.PSObject.Properties[$env:REMOTE_FILE]; $helperEntry=$m.files.PSObject.Properties[$env:REMOTE_HELPER_FILE]; $committerEntry=$m.files.PSObject.Properties[$env:REMOTE_COMMITTER_FILE]; if($null -eq $batEntry -or $null -eq $helperEntry -or $null -eq $committerEntry -or @($m.files.PSObject.Properties).Count -ne 3){exit 32}; $batExpected=([string]$batEntry.Value).ToLowerInvariant(); $helperExpected=([string]$helperEntry.Value).ToLowerInvariant(); $committerExpected=([string]$committerEntry.Value).ToLowerInvariant(); if($batExpected -notmatch '^[0-9a-f]{64}$' -or $helperExpected -notmatch '^[0-9a-f]{64}$' -or $committerExpected -notmatch '^[0-9a-f]{64}$'){exit 33}; $batInfo=Get-Item -LiteralPath $env:UPDATE_BAT_PATH; $helperInfo=Get-Item -LiteralPath $env:UPDATE_HELPER_PATH; $committerInfo=Get-Item -LiteralPath $env:UPDATE_COMMITTER_PATH; if($batInfo.Length -lt 30000 -or $batInfo.Length -gt 500000 -or $helperInfo.Length -lt 4000 -or $helperInfo.Length -gt 200000 -or $committerInfo.Length -lt 8000 -or $committerInfo.Length -gt 250000){exit 34}; $batText=[IO.File]::ReadAllText($env:UPDATE_BAT_PATH); $vm=[regex]::Matches($batText,'(?mi)^\s*set\s+"LOCAL_VER=([0-9]+)"\s*$'); if($vm.Count -ne 1 -or [int64]$vm[0].Groups[1].Value -ne [int64]$env:REMOTE_VER){exit 35}; $batRequired=@('@echo off','update-channel.json','usbdiag.shared.local.cmd','wifi.local.cmd','usbdiag-wifi-readiness.ps1','usbdiag-bundle-commit.ps1','USB_DIAG_PIPELINE_ACCEPTED_V1','CITYPC_USB_TOKEN','Invoke-RestMethod',':CONFIRM_INSTALLED_BUNDLE',':LAUNCH_UPDATE_RECOVERY',':INICIO',('set "REMOTE_FILE=' + $env:REMOTE_FILE + '"')); foreach($needle in $batRequired){if($batText.IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -lt 0){exit 36}}; $helperText=[IO.File]::ReadAllText($env:UPDATE_HELPER_PATH); $helperRequired=@('param(','Invoke-NetshBounded','Assert-TlsHandshake','Remove-DiagnosticWifiProfile','Connect','Cleanup'); foreach($needle in $helperRequired){if($helperText.IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -lt 0){exit 37}}; $committerText=[IO.File]::ReadAllText($env:UPDATE_COMMITTER_PATH); $committerRequired=@("ValidateSet('Commit', 'Confirm', 'Recover')",'Wait-ForParentExit','Invoke-Confirm','Invoke-Recover','Restore-FromState','rollbackConfirmed','private configuration is not updateable','installed-awaiting-confirmation'); foreach($needle in $committerRequired){if($committerText.IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -lt 0){exit 38}}; $batActual=(Get-FileHash -LiteralPath $env:UPDATE_BAT_PATH -Algorithm SHA256).Hash.ToLowerInvariant(); $helperActual=(Get-FileHash -LiteralPath $env:UPDATE_HELPER_PATH -Algorithm SHA256).Hash.ToLowerInvariant(); $committerActual=(Get-FileHash -LiteralPath $env:UPDATE_COMMITTER_PATH -Algorithm SHA256).Hash.ToLowerInvariant(); if($batActual -cne $batExpected -or $helperActual -cne $helperExpected -or $committerActual -cne $committerExpected){exit 39}; exit 0 } catch { exit 40 }" >nul 2>&1
if errorlevel 1 goto :UPDATE_INVALID

if "!VERSION_CHECK!"=="20" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$m=Get-Content -LiteralPath $env:UPDATE_MANIFEST_PATH -Raw -Encoding UTF8 ^| ConvertFrom-Json; $items=@(@($env:REMOTE_FILE,$env:UPDATE_BAT_PATH),@($env:REMOTE_HELPER_FILE,$env:UPDATE_HELPER_PATH),@($env:REMOTE_COMMITTER_FILE,$env:UPDATE_COMMITTER_PATH)); foreach($item in $items){$installed=[IO.Path]::Combine($env:SCRIPT_DIR,$item[0]); if(-not (Test-Path -LiteralPath $installed -PathType Leaf)){exit 1}; $expected=([string]$m.files.PSObject.Properties[$item[0]].Value).ToLowerInvariant(); if((Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expected){exit 1}}; exit 0" >nul 2>&1
    if not errorlevel 1 (
        echo [OK] Version V%LOCAL_VER% y hashes del bundle estan vigentes.
        echo.
        call :CLEAN_UPDATE_FILES
        goto :skip_update_diag
    )
    echo [AVISO] Version V%LOCAL_VER% con bundle corrupto o incompleto. Reparando...
) else (
    echo [!!] Nueva version disponible: V!REMOTE_VER! ^(actual: V%LOCAL_VER%^)
)

:: El committer temporal espera que este BAT salga y aplica el bundle completo.
:: Limpia Wi-Fi, pero conserva el lock hasta que el committer termine.
if defined WIFI_READY (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%WIFI_HELPER%" -Action Cleanup >nul 2>&1
    if errorlevel 1 goto :ERROR_WIFI_CLEANUP
    set "WIFI_READY="
)
if defined WIFI_XML if exist "%WIFI_XML%" del /F /Q "%WIFI_XML%" >nul 2>&1

echo.
echo [INFO] Bundle V!REMOTE_VER! validado. Aplicando BAT, helper y committer...
set "UPDATE_TRANSACTION_TOKEN="
for /f "usebackq delims=" %%T in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "[Console]::Write([guid]::NewGuid().ToString('N'))"`) do set "UPDATE_TRANSACTION_TOKEN=%%T"
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:UPDATE_TRANSACTION_TOKEN -cmatch '^[0-9a-f]{32}$'){exit 0}else{exit 1}" >nul 2>&1
if errorlevel 1 goto :UPDATE_COMMIT_START_FAILED
call :CAPTURE_PARENT_PID
if errorlevel 1 goto :UPDATE_COMMIT_START_FAILED
start "" powershell -NoProfile -ExecutionPolicy Bypass -File "!UPDATE_COMMITTER_PATH!" -Action Commit -ScriptDir "%SCRIPT_DIR%" -TransactionDir "%UPDATE_TRANSACTION_DIR%" -LockDir "%INSTANCE_LOCK%" -ParentPid "!UPDATE_PARENT_PID!" -Token "!UPDATE_TRANSACTION_TOKEN!" -RemoteVersion "!REMOTE_VER!" -BatName "!REMOTE_FILE!" -HelperName "!REMOTE_HELPER_FILE!" -CommitterName "!REMOTE_COMMITTER_FILE!" -DownloadedBat "!UPDATE_BAT_PATH!" -DownloadedHelper "!UPDATE_HELPER_PATH!" -DownloadedCommitter "!UPDATE_COMMITTER_PATH!" -ManifestPath "!UPDATE_MANIFEST_PATH!" -ManifestExpectedSha256 "!EXPECTED_MANIFEST_HASH!" -ChannelPath "!UPDATE_CHANNEL_PATH!" -ValidatedVersionPath "!UPDATE_VALID_VERSION_PATH!" -CommitPath "!UPDATE_COMMIT_PATH!" -ManifestExpectedHashPath "!UPDATE_MANIFEST_EXPECTED_HASH_PATH!"
if errorlevel 1 goto :UPDATE_COMMIT_START_FAILED
set "INSTANCE_LOCK_HELD="
exit

:UPDATE_COMMIT_START_FAILED
echo [ERROR] No se pudo iniciar el committer atomico. Se conserva V%LOCAL_VER%.
echo.
call :CLEAN_UPDATE_FILES
call :RELEASE_INSTANCE_LOCK
pause
exit /b 34

:UPDATE_INVALID
echo [ERROR] La actualizacion no paso version, estructura o SHA-256. Usando V%LOCAL_VER%.
echo.
call :CLEAN_UPDATE_FILES
goto :skip_update_diag

:UPDATE_UNAVAILABLE
echo [AVISO] No se pudo verificar o descargar la actualizacion. Usando V%LOCAL_VER%.
echo.
call :CLEAN_UPDATE_FILES
goto :skip_update_diag

:skip_update_diag


:INICIO
cls
echo ==========================================
echo      SISTEMA DE DIAGNOSTICO CITYPC (V%LOCAL_VER%)
echo ==========================================
echo.
set "TICKET="
setlocal DisableDelayedExpansion
set /p "TICKET=Ingrese # Ticket (5 digitos): "
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:TICKET -cmatch '^[0-9]{5}$'){exit 0}else{exit 1}" >nul 2>&1
if errorlevel 1 (
    endlocal
    goto ERROR_TICKET
)
set "TICKET_VALID=%TICKET%"
endlocal & set "TICKET=%TICKET_VALID%"

:TIPO_EQUIPO
echo.
echo Tipo de equipo:
echo   1. Laptop
echo   2. CPU (Escritorio)
echo   3. AIO (All-in-One)
echo.
set "TIPO="
set "TIPO_NOMBRE="
set /p "TIPO=Seleccione (1, 2 o 3): "

if "!TIPO!"=="1" set "TIPO_NOMBRE=Laptop"
if "!TIPO!"=="2" set "TIPO_NOMBRE=CPU"
if "!TIPO!"=="3" set "TIPO_NOMBRE=AIO"

if not defined TIPO_NOMBRE (
    echo.
    echo [ERROR] Opcion invalida. Seleccione 1, 2 o 3.
    set "TIPO_NOMBRE="
    goto TIPO_EQUIPO
)

echo.
echo [OK] Equipo: !TIPO_NOMBRE! - Ticket: %TICKET%

echo.
echo [1/8] Preparando herramientas...
set "psfile=%temp%\citypc_diag_v8_!RANDOM!!RANDOM!.ps1"
if exist "%psfile%" del "%psfile%"

:: ---------------------------------------------------------
:: 3. GENERANDO SCRIPT POWERSHELL
:: ---------------------------------------------------------
echo $ErrorActionPreference = 'SilentlyContinue' >> "%psfile%"
echo [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 >> "%psfile%"
echo $ticket = "%TICKET%" >> "%psfile%"
echo $tipoEquipo = "%TIPO_NOMBRE%" >> "%psfile%"
echo $webhookUrl = "%WEBHOOK_URL%" >> "%psfile%"
echo $path = [Environment]::GetFolderPath('Desktop') >> "%psfile%"
echo $archivo = "$path\Reporte_Ticket_$ticket.txt" >> "%psfile%"

echo function Log-Dual($txt, $col="White") { >> "%psfile%"
echo     Write-Host $txt -ForegroundColor $col >> "%psfile%"
echo     $txt ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo } >> "%psfile%"

echo Add-Type -AssemblyName System.Windows.Forms >> "%psfile%"
echo Add-Type -AssemblyName System.Drawing >> "%psfile%"

:: --- ENCABEZADO ---
echo "==========================================" ^| Out-File -FilePath $archivo -Encoding UTF8 >> "%psfile%"
echo " DIAGNOSTICO DE SISTEMA - TICKET $ticket" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo " TIPO DE EQUIPO: $tipoEquipo" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "==========================================" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 1: SISTEMA Y PROCESADOR (ampliado con MB, OS, build)
:: -------------------------------------------------------------------
echo $sys  = Get-CimInstance Win32_ComputerSystem >> "%psfile%"
echo $bios = Get-CimInstance Win32_BIOS >> "%psfile%"
echo $cpu  = Get-CimInstance Win32_Processor >> "%psfile%"
echo $mb   = Get-CimInstance Win32_BaseBoard >> "%psfile%"
echo $os   = Get-CimInstance Win32_OperatingSystem >> "%psfile%"
echo $ghz  = [math]::Round($cpu.MaxClockSpeed / 1000, 2) >> "%psfile%"

echo Clear-Host >> "%psfile%"
echo Log-Dual "=== 1. SISTEMA Y PROCESADOR ===" "Cyan" >> "%psfile%"
echo Log-Dual "   * Fabricante:    $($sys.Manufacturer)" >> "%psfile%"
echo Log-Dual "   * Modelo:        $($sys.Model)" >> "%psfile%"
echo Log-Dual "   * Serial BIOS:   $($bios.SerialNumber)" >> "%psfile%"
echo Log-Dual "   * Motherboard:   $($mb.Manufacturer) $($mb.Product)" >> "%psfile%"
echo Log-Dual "   * Serial MB:     $($mb.SerialNumber)" >> "%psfile%"
echo Log-Dual "   * CPU:           $($cpu.Name)" >> "%psfile%"
echo Log-Dual "   * Velocidad:     $ghz GHz  /  $($cpu.NumberOfCores) Nucleos  /  $($cpu.NumberOfLogicalProcessors) Hilos" >> "%psfile%"
echo Log-Dual "   * OS:            $($os.Caption) $($os.OSArchitecture)" >> "%psfile%"
echo Log-Dual "   * Build:         $($os.BuildNumber)" >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 2: ALMACENAMIENTO + SALUD SSD/HDD
:: -------------------------------------------------------------------
echo Log-Dual "=== 2. ALMACENAMIENTO ===" "Cyan" >> "%psfile%"
echo $allDisks = Get-PhysicalDisk ^| Where-Object { $_.BusType -ne 'USB' } ^| Sort-Object MediaType >> "%psfile%"
echo if ($allDisks.Count -eq 0) { >> "%psfile%"
echo     Log-Dual "   * Sin discos internos detectados" "Yellow" >> "%psfile%"
echo } >> "%psfile%"
echo foreach ($d in $allDisks) { >> "%psfile%"
echo     $gb = [math]::Round($d.Size/1GB) >> "%psfile%"
echo     $healthColor = if ($d.HealthStatus -eq 'Healthy') {"Green"} elseif ($d.HealthStatus -eq 'Warning') {"Yellow"} else {"Red"} >> "%psfile%"
echo     Log-Dual "   * Disco:         $($d.FriendlyName)  [$($d.BusType)]" >> "%psfile%"
echo     Log-Dual "     Tipo/Tam:      $($d.MediaType) - $gb GB" >> "%psfile%"
echo     Log-Dual "     Salud/Estado:  $($d.HealthStatus) - $($d.OperationalStatus)" $healthColor >> "%psfile%"
echo     $rel = Get-StorageReliabilityCounter -PhysicalDisk $d -ErrorAction SilentlyContinue >> "%psfile%"
echo     if ($rel) { >> "%psfile%"
echo         if ($null -ne $rel.Wear -and $rel.Wear -gt 0) { >> "%psfile%"
echo             $wearCol = if ($rel.Wear -gt 80) {"Red"} elseif ($rel.Wear -gt 50) {"Yellow"} else {"Green"} >> "%psfile%"
echo             Log-Dual "     Desgaste:      $($rel.Wear)%% de vida usada" $wearCol >> "%psfile%"
echo         } >> "%psfile%"
echo         if ($null -ne $rel.Temperature -and $rel.Temperature -gt 0) { >> "%psfile%"
echo             $tempCol = if ($rel.Temperature -gt 55) {"Red"} elseif ($rel.Temperature -gt 45) {"Yellow"} else {"Green"} >> "%psfile%"
echo             Log-Dual "     Temperatura:   $($rel.Temperature) C" $tempCol >> "%psfile%"
echo         } >> "%psfile%"
echo         if ($null -ne $rel.ReadErrorsTotal -and $rel.ReadErrorsTotal -gt 0) { >> "%psfile%"
echo             Log-Dual "     Err.Lectura:   $($rel.ReadErrorsTotal)  ADVERTENCIA" "Red" >> "%psfile%"
echo         } >> "%psfile%"
echo         if ($null -ne $rel.WriteErrorsTotal -and $rel.WriteErrorsTotal -gt 0) { >> "%psfile%"
echo             Log-Dual "     Err.Escritura: $($rel.WriteErrorsTotal)  ADVERTENCIA" "Red" >> "%psfile%"
echo         } >> "%psfile%"
echo     } >> "%psfile%"
echo } >> "%psfile%"

echo $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue >> "%psfile%"
echo if ($bl) { >> "%psfile%"
echo     if ($bl.ProtectionStatus -eq 'On') { $blStatus = "ACTIVADO (RIESGO SI NO HAY CLAVE)" } else { $blStatus = "DESACTIVADO (LIBRE)" } >> "%psfile%"
echo } else { $blStatus = "No disponible / Windows Home" } >> "%psfile%"
echo Log-Dual "   * BitLocker:     $blStatus" "Yellow" >> "%psfile%"

echo $vol = Get-Volume ^| Where-Object {$_.DriveLetter -eq 'C'} >> "%psfile%"
echo if ($vol) { >> "%psfile%"
echo     $libre   = [math]::Round($vol.SizeRemaining/1GB, 1) >> "%psfile%"
echo     $totalVol = [math]::Round($vol.Size/1GB, 1) >> "%psfile%"
echo     $freeCol = if ($libre -lt 10) {"Red"} elseif ($libre -lt 30) {"Yellow"} else {"Green"} >> "%psfile%"
echo     Log-Dual "   * Unidad C:      $libre GB libres de $totalVol GB" $freeCol >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 3: MEMORIA RAM
:: -------------------------------------------------------------------
echo $mem   = @(Get-CimInstance Win32_PhysicalMemory ^| Sort-Object DeviceLocator, BankLabel) >> "%psfile%"
echo $slots = Get-CimInstance Win32_PhysicalMemoryArray >> "%psfile%"
echo Log-Dual "=== 3. MEMORIA RAM ===" "Cyan" >> "%psfile%"
echo $totalMem   = [math]::Round(($mem ^| Measure-Object -Property Capacity -Sum).Sum / 1GB) >> "%psfile%"
echo $usados     = @($mem).Count >> "%psfile%"
echo $libresSlot = $slots.MemoryDevices - $usados >> "%psfile%"
echo Log-Dual "   * Instalada:     $totalMem GB  ($($mem[0].Speed) MHz)" >> "%psfile%"
echo Log-Dual "   * Ranuras:       $usados Ocupadas / $libresSlot Libres" >> "%psfile%"
echo $ramTypeMap = @{20='DDR';21='DDR2';22='DDR2 FB-DIMM';24='DDR3';26='DDR4';27='LPDDR';28='LPDDR2';29='LPDDR3';30='LPDDR4';34='DDR5';35='LPDDR5'} >> "%psfile%"
echo $badSerials = @('UNKNOWN','NONE','N/A','NA','NOT SPECIFIED','NOT AVAILABLE','TO BE FILLED BY O.E.M.','TO BE FILLED BY OEM','DEFAULT STRING') >> "%psfile%"
echo $moduleIndex = 0 >> "%psfile%"
echo foreach ($module in $mem) { >> "%psfile%"
echo     $moduleIndex++ >> "%psfile%"
echo     $moduleGB = [math]::Round([double]$module.Capacity / 1GB, 2) >> "%psfile%"
echo     $typeCode = [int]$module.SMBIOSMemoryType >> "%psfile%"
echo     if ($typeCode -eq 0) { $typeCode = [int]$module.MemoryType } >> "%psfile%"
echo     $moduleType = if ($ramTypeMap.ContainsKey($typeCode)) { $ramTypeMap[$typeCode] } else { 'Desconocido' } >> "%psfile%"
echo     $serial = ([string]$module.SerialNumber).Trim() >> "%psfile%"
echo     $serialUpper = $serial.ToUpperInvariant() >> "%psfile%"
echo     $serialCompact = ($serial -replace '[\s-]', '').ToUpperInvariant() >> "%psfile%"
echo     if ([string]::IsNullOrWhiteSpace($serial) -or $badSerials -contains $serialUpper -or $serialCompact -match '^0+$' -or $serialCompact -match '^F+$') { $serial = 'No disponible' } >> "%psfile%"
echo     Log-Dual ('   * Modulo {0}: {1} GB; Tipo: {2}; Serial: {3}' -f $moduleIndex, $moduleGB, $moduleType, $serial) >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 4: TARJETA GRAFICA (NUEVA)
:: -------------------------------------------------------------------
echo Log-Dual "=== 4. TARJETA GRAFICA ===" "Cyan" >> "%psfile%"
echo $gpus = Get-CimInstance Win32_VideoController >> "%psfile%"
echo foreach ($gpu in $gpus) { >> "%psfile%"
echo     if (-not $gpu.Name) { continue } >> "%psfile%"
echo     $vramBytes = [long]$gpu.AdapterRAM >> "%psfile%"
echo     if ($vramBytes -gt 0) { >> "%psfile%"
echo         $vramMB = [math]::Round($vramBytes / 1MB) >> "%psfile%"
echo         $vramDisplay = if ($vramMB -ge 1024) { "$([math]::Round($vramMB/1024,1)) GB" } else { "$vramMB MB" } >> "%psfile%"
echo     } else { $vramDisplay = "Ver Administrador de Dispositivos" } >> "%psfile%"
echo     $gpuStatusCol = if ($gpu.Status -eq 'OK') {"Green"} else {"Red"} >> "%psfile%"
echo     Log-Dual "   * GPU:           $($gpu.Name)" >> "%psfile%"
echo     Log-Dual "     VRAM:          $vramDisplay" >> "%psfile%"
echo     Log-Dual "     Driver:        $($gpu.DriverVersion)" >> "%psfile%"
echo     Log-Dual "     Resolucion:    $($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution)" >> "%psfile%"
echo     Log-Dual "     Estado:        $($gpu.Status)" $gpuStatusCol >> "%psfile%"
echo     Log-Dual "" >> "%psfile%"
echo } >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 5: BATERIA (NUEVA)
:: -------------------------------------------------------------------
echo Log-Dual "=== 5. BATERIA ===" "Cyan" >> "%psfile%"
echo $bat = Get-CimInstance Win32_Battery >> "%psfile%"
echo if ($bat) { >> "%psfile%"
echo     $charge = $bat.EstimatedChargeRemaining >> "%psfile%"
echo     $batMap = @{1='Descargando (Sin CA)';2='Con Corriente CA';3='Totalmente Cargada';4='Carga Baja';5='Carga Critica';6='Cargando';7='Cargando (Alta)';8='Cargando (Baja)';9='Cargando (Critica)';11='Sin Red CA'} >> "%psfile%"
echo     $batStatusText = if ($batMap.ContainsKey([int]$bat.BatteryStatus)) { $batMap[[int]$bat.BatteryStatus] } else { "Desconocido" } >> "%psfile%"
echo     $batCol = if ($charge -lt 20) {"Red"} elseif ($charge -lt 50) {"Yellow"} else {"Green"} >> "%psfile%"
echo     Log-Dual "   * Carga:         $charge%%" $batCol >> "%psfile%"
echo     Log-Dual "   * Estado:        $batStatusText" >> "%psfile%"
echo     Log-Dual "   * Condicion:     $($bat.Status)" >> "%psfile%"
echo     if ($bat.EstimatedRunTime -and $bat.EstimatedRunTime -lt 900000) { >> "%psfile%"
echo         Log-Dual "   * Autonomia:     $($bat.EstimatedRunTime) min aprox" >> "%psfile%"
echo     } >> "%psfile%"
echo } else { >> "%psfile%"
echo     Log-Dual "   * Bateria:       No detectada (posiblemente Escritorio)" "Yellow" >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: FUNCION MOSTRAR-PREGUNTA (sin cambios)
:: -------------------------------------------------------------------
echo function Mostrar-Pregunta($colorFondo, $colorTexto, $pregunta, $txtBtnOK="TODO BIEN", $txtBtnFail="CON DEFECTOS") { >> "%psfile%"
echo     $form = New-Object System.Windows.Forms.Form >> "%psfile%"
echo     $form.FormBorderStyle = 'None' >> "%psfile%"
echo     $form.WindowState = 'Maximized' >> "%psfile%"
echo     $form.BackColor = $colorFondo >> "%psfile%"
echo     $form.TopMost = $true >> "%psfile%"
echo     $res = "Sin Probar" >> "%psfile%"
echo     $lbl = New-Object System.Windows.Forms.Label >> "%psfile%"
echo     $lbl.Text = $pregunta >> "%psfile%"
echo     $lbl.ForeColor = $colorTexto >> "%psfile%"
echo     $lbl.BackColor = [System.Drawing.Color]::Transparent >> "%psfile%"
echo     $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo     $lbl.AutoSize = $true >> "%psfile%"
echo     $lbl.TextAlign = 'MiddleCenter' >> "%psfile%"
echo     $lbl.Location = New-Object System.Drawing.Point(300, 250) >> "%psfile%"
echo     $form.Controls.Add($lbl) >> "%psfile%"
echo     $btnOk = New-Object System.Windows.Forms.Button >> "%psfile%"
echo     $btnOk.Text = $txtBtnOK >> "%psfile%"
echo     $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 16) >> "%psfile%"
echo     $btnOk.Size = New-Object System.Drawing.Size(250, 80) >> "%psfile%"
echo     $btnOk.Location = New-Object System.Drawing.Point(400, 450) >> "%psfile%"
echo     $btnOk.BackColor = 'White' >> "%psfile%"
echo     $btnOk.ForeColor = 'Black' >> "%psfile%"
echo     $btnOk.Add_Click({ $script:res = "OK"; $form.Close() }) >> "%psfile%"
echo     $form.Controls.Add($btnOk) >> "%psfile%"
echo     $btnFail = New-Object System.Windows.Forms.Button >> "%psfile%"
echo     $btnFail.Text = $txtBtnFail >> "%psfile%"
echo     $btnFail.Font = New-Object System.Drawing.Font("Segoe UI", 16) >> "%psfile%"
echo     $btnFail.Size = New-Object System.Drawing.Size(250, 80) >> "%psfile%"
echo     $btnFail.Location = New-Object System.Drawing.Point(700, 450) >> "%psfile%"
echo     $btnFail.BackColor = 'White' >> "%psfile%"
echo     $btnFail.ForeColor = 'Red' >> "%psfile%"
echo     $btnFail.Add_Click({ $script:res = "CON FALLAS"; $form.Close() }) >> "%psfile%"
echo     $form.Controls.Add($btnFail) >> "%psfile%"
echo     $form.ShowDialog() ^| Out-Null >> "%psfile%"
echo     return $script:res >> "%psfile%"
echo } >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 6: DIAGNOSTICO DE PANTALLA (solo Laptop y AIO)
:: -------------------------------------------------------------------
echo if ($tipoEquipo -ne 'CPU') { >> "%psfile%"
echo Log-Dual "=== 6. DIAGNOSTICO DE PANTALLA ===" "Magenta" >> "%psfile%"
echo Start-Sleep -Seconds 1 >> "%psfile%"
echo $p1 = Mostrar-Pregunta "White" "Black" "Fondo BLANCO. Ves manchas oscuras?" "PANTALLA LIMPIA" "TIENE MANCHAS" >> "%psfile%"
echo Log-Dual "   * Blancos:       $p1" "White" >> "%psfile%"
echo Start-Sleep -Seconds 1 >> "%psfile%"
echo $p2 = Mostrar-Pregunta "Black" "White" "Fondo NEGRO. Ves pixeles muertos o luz?" "PANTALLA PERFECTA" "TIENE DEFECTOS" >> "%psfile%"
echo Log-Dual "   * Negros:        $p2" "White" >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 7: DIAGNOSTICO DE AUDIO
:: -------------------------------------------------------------------
echo Log-Dual "=== 7. DIAGNOSTICO DE AUDIO ===" "Yellow" >> "%psfile%"
echo $code = @' >> "%psfile%"
echo using System; using System.Runtime.InteropServices; using System.Text; >> "%psfile%"
echo public class Audio { >> "%psfile%"
echo   [DllImport("winmm.dll", EntryPoint="mciSendStringA")] >> "%psfile%"
echo   public static extern int mci(string cmd, StringBuilder ret, int len, IntPtr h); >> "%psfile%"
echo } >> "%psfile%"
echo '@ >> "%psfile%"
echo Add-Type -TypeDefinition $code -PassThru ^| Out-Null >> "%psfile%"

echo Write-Host "   Emitiendo tonos..." >> "%psfile%"
echo [Console]::Beep(3000, 300); Start-Sleep -m 150 >> "%psfile%"
echo [Console]::Beep(1000, 300); Start-Sleep -m 150 >> "%psfile%"
echo [Console]::Beep(500, 500) >> "%psfile%"
echo $pa = Mostrar-Pregunta "DarkBlue" "White" "Sonaron los 3 tonos correctamente?" "SONIDO OK" "FALLA DE SONIDO" >> "%psfile%"
echo Log-Dual "   * Bocinas:       $pa" "White" >> "%psfile%"

echo Write-Host "   Probando Mic (HABLE AHORA)..." -ForegroundColor Red -BackgroundColor White >> "%psfile%"
echo [Audio]::mci("open new type waveaudio alias rec", $null, 0, 0) >> "%psfile%"
echo [Audio]::mci("record rec", $null, 0, 0) >> "%psfile%"
echo Start-Sleep -Seconds 4 >> "%psfile%"
echo [Audio]::mci("save rec $env:temp\test_mic.wav", $null, 0, 0) >> "%psfile%"
echo [Audio]::mci("close rec", $null, 0, 0) >> "%psfile%"
echo (New-Object Media.SoundPlayer "$env:temp\test_mic.wav").PlaySync() >> "%psfile%"
echo $pm = Mostrar-Pregunta "DarkBlue" "White" "Escuchaste la grabacion de voz?" "MICROFONO OK" "FALLA MICROFONO" >> "%psfile%"
echo Log-Dual "   * Microfono:     $pm" "White" >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: SECCION 8: PRUEBA DE TECLADO (solo Laptop)
:: -------------------------------------------------------------------
echo } >> "%psfile%"
echo if ($tipoEquipo -eq 'Laptop') { >> "%psfile%"
echo Log-Dual "=== 8. PRUEBA DE TECLADO ===" "Green" >> "%psfile%"

:: ------ DIALOGO PREVIO: solo pregunta numpad ------
echo $preForm = New-Object System.Windows.Forms.Form >> "%psfile%"
echo $preForm.FormBorderStyle = 'None' >> "%psfile%"
echo $preForm.Size = New-Object System.Drawing.Size(680, 260) >> "%psfile%"
echo $preForm.StartPosition = 'CenterScreen' >> "%psfile%"
echo $preForm.BackColor = [System.Drawing.Color]::FromArgb(18,18,18) >> "%psfile%"
echo $preForm.TopMost = $true >> "%psfile%"
echo $preForm.AcceptButton = $null >> "%psfile%"

echo $preLbl1 = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $preLbl1.Text = "PRUEBA DE TECLADO - ESPANOL" >> "%psfile%"
echo $preLbl1.ForeColor = [System.Drawing.Color]::Cyan >> "%psfile%"
echo $preLbl1.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $preLbl1.AutoSize = $true >> "%psfile%"
echo $preLbl1.Location = New-Object System.Drawing.Point(30, 20) >> "%psfile%"
echo $preForm.Controls.Add($preLbl1) >> "%psfile%"

echo $preLbl2 = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $preLbl2.Text = "El teclado tiene bloque numerico (numpad) a la derecha?" >> "%psfile%"
echo $preLbl2.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $preLbl2.Font = New-Object System.Drawing.Font("Segoe UI", 14) >> "%psfile%"
echo $preLbl2.AutoSize = $true >> "%psfile%"
echo $preLbl2.Location = New-Object System.Drawing.Point(30, 75) >> "%psfile%"
echo $preForm.Controls.Add($preLbl2) >> "%psfile%"

echo $script:kbNumpad = $false >> "%psfile%"

echo $btnNumSi = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $btnNumSi.Text = "SI - Tiene Numpad" >> "%psfile%"
echo $btnNumSi.Font = New-Object System.Drawing.Font("Segoe UI", 13) >> "%psfile%"
echo $btnNumSi.Size = New-Object System.Drawing.Size(200, 70) >> "%psfile%"
echo $btnNumSi.Location = New-Object System.Drawing.Point(30, 140) >> "%psfile%"
echo $btnNumSi.BackColor = [System.Drawing.Color]::FromArgb(50,50,50) >> "%psfile%"
echo $btnNumSi.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $btnNumSi.Add_Click({ $script:kbNumpad = $true; $btnNumSi.BackColor = [System.Drawing.Color]::FromArgb(0,160,0); $btnNumNo.BackColor = [System.Drawing.Color]::FromArgb(50,50,50) }) >> "%psfile%"
echo $preForm.Controls.Add($btnNumSi) >> "%psfile%"

echo $btnNumNo = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $btnNumNo.Text = "NO - Sin Numpad" >> "%psfile%"
echo $btnNumNo.Font = New-Object System.Drawing.Font("Segoe UI", 13) >> "%psfile%"
echo $btnNumNo.Size = New-Object System.Drawing.Size(200, 70) >> "%psfile%"
echo $btnNumNo.Location = New-Object System.Drawing.Point(250, 140) >> "%psfile%"
echo $btnNumNo.BackColor = [System.Drawing.Color]::FromArgb(0,100,160) >> "%psfile%"
echo $btnNumNo.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $btnNumNo.Add_Click({ $script:kbNumpad = $false; $btnNumNo.BackColor = [System.Drawing.Color]::FromArgb(0,160,0); $btnNumSi.BackColor = [System.Drawing.Color]::FromArgb(50,50,50) }) >> "%psfile%"
echo $preForm.Controls.Add($btnNumNo) >> "%psfile%"

echo $btnPreOk = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $btnPreOk.Text = "INICIAR" >> "%psfile%"
echo $btnPreOk.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $btnPreOk.Size = New-Object System.Drawing.Size(160, 70) >> "%psfile%"
echo $btnPreOk.Location = New-Object System.Drawing.Point(490, 140) >> "%psfile%"
echo $btnPreOk.BackColor = [System.Drawing.Color]::FromArgb(180,100,0) >> "%psfile%"
echo $btnPreOk.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $btnPreOk.Add_Click({ $preForm.Close() }) >> "%psfile%"
echo $preForm.Controls.Add($btnPreOk) >> "%psfile%"

echo $preForm.ShowDialog() ^| Out-Null >> "%psfile%"

:: ------ DEFINIR FILAS: pares [KeyCode, Texto] ------
:: Enye se construye en runtime para conservar el BAT en ASCII
echo $kbEnyeName = [string][char]0x00D1 >> "%psfile%"
echo $rowNums = @(@('D1','1'),@('D2','2'),@('D3','3'),@('D4','4'),@('D5','5'),@('D6','6'),@('D7','7'),@('D8','8'),@('D9','9'),@('D0','0'),@('Back','BKSP')) >> "%psfile%"
echo $rowQ = @(@('Q','Q'),@('W','W'),@('E','E'),@('R','R'),@('T','T'),@('Y','Y'),@('U','U'),@('I','I'),@('O','O'),@('P','P')) >> "%psfile%"
echo $rowA = @(@('A','A'),@('S','S'),@('D','D'),@('F','F'),@('G','G'),@('H','H'),@('J','J'),@('K','K'),@('L','L'),@('Enye',$kbEnyeName)) >> "%psfile%"
echo $rowZ = @(@('Z','Z'),@('X','X'),@('C','C'),@('V','V'),@('B','B'),@('N','N'),@('M','M')) >> "%psfile%"
echo $kbMainRows = @($rowNums, $rowQ, $rowA, $rowZ) >> "%psfile%"

:: Numpad: solo NumPad0-NumPad9; Insert no cuenta como NumPad0
echo $rowNP1 = @(@('NumPad7','N7'),@('NumPad8','N8'),@('NumPad9','N9')) >> "%psfile%"
echo $rowNP2 = @(@('NumPad4','N4'),@('NumPad5','N5'),@('NumPad6','N6')) >> "%psfile%"
echo $rowNP3 = @(@('NumPad1','N1'),@('NumPad2','N2'),@('NumPad3','N3')) >> "%psfile%"
echo $rowNP4 = @(,@('NumPad0','N0')) >> "%psfile%"
echo $kbNumRows = @($rowNP1,$rowNP2,$rowNP3,$rowNP4) >> "%psfile%"

:: 38 objetivos sin numpad; 48 con numpad
echo $kbAllCodes = @('D1','D2','D3','D4','D5','D6','D7','D8','D9','D0','Back','Q','W','E','R','T','Y','U','I','O','P','A','S','D','F','G','H','J','K','L','Enye','Z','X','C','V','B','N','M') >> "%psfile%"
echo $kbAllNames = @('1','2','3','4','5','6','7','8','9','0','BKSP','Q','W','E','R','T','Y','U','I','O','P','A','S','D','F','G','H','J','K','L',$kbEnyeName,'Z','X','C','V','B','N','M') >> "%psfile%"
echo if ($script:kbNumpad) { >> "%psfile%"
echo     $kbAllCodes += @('NumPad7','NumPad8','NumPad9','NumPad4','NumPad5','NumPad6','NumPad1','NumPad2','NumPad3','NumPad0') >> "%psfile%"
echo     $kbAllNames += @('N7','N8','N9','N4','N5','N6','N1','N2','N3','N0') >> "%psfile%"
echo } >> "%psfile%"

echo $script:kbOutcome = 'NoCompletada' >> "%psfile%"
echo $script:kbConfirmPending = $false >> "%psfile%"
echo $script:kbAllowClose = $false >> "%psfile%"
echo $script:kbMissingFinal = @() >> "%psfile%"

:: ------ VENTANA PRINCIPAL ------
echo $kbForm = New-Object System.Windows.Forms.Form >> "%psfile%"
echo $kbForm.FormBorderStyle = 'None' >> "%psfile%"
echo $kbForm.WindowState = 'Maximized' >> "%psfile%"
echo $kbForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None >> "%psfile%"
echo $kbForm.BackColor = [System.Drawing.Color]::FromArgb(18,18,18) >> "%psfile%"
echo $kbForm.TopMost = $true >> "%psfile%"
echo $kbForm.KeyPreview = $true >> "%psfile%"
echo $kbForm.AcceptButton = $null >> "%psfile%"
echo $kbForm.CancelButton = $null >> "%psfile%"

echo $npAviso = if ($script:kbNumpad) {'  NUMPAD: active NumLock antes de probar'} else {''} >> "%psfile%"
echo $kbTitle = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $kbTitle.Text = "PRUEBA DE TECLADO - ESPANOL  --  Presione cada tecla  --  Enter: REVISAR / FINALIZAR$npAviso" >> "%psfile%"
echo $kbTitle.ForeColor = [System.Drawing.Color]::Cyan >> "%psfile%"
echo $kbTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $kbTitle.AutoSize = $true >> "%psfile%"
echo $kbTitle.Location = New-Object System.Drawing.Point(20, 12) >> "%psfile%"
echo $kbForm.Controls.Add($kbTitle) >> "%psfile%"

echo $kbLabels  = @{} >> "%psfile%"
echo $kbPressed = [System.Collections.Generic.HashSet[string]]::new() >> "%psfile%"
echo if ($script:kbNumpad) { $kbKeyW = 64; $kbKeyH = 64; $kbGap = 4 } else { $kbKeyW = 76; $kbKeyH = 70; $kbGap = 5 } >> "%psfile%"
echo $kbStartY = 55 >> "%psfile%"

:: Dibujar teclas principales en un area compatible con 1280x720
echo foreach ($row in $kbMainRows) { >> "%psfile%"
echo     $kbX = if ($script:kbNumpad) {24} else {40} >> "%psfile%"
echo     foreach ($pair in $row) { >> "%psfile%"
echo         $kcode = $pair[0]; $ktxt = $pair[1] >> "%psfile%"
echo         $lbl2 = New-Object System.Windows.Forms.Label >> "%psfile%"
echo         $lbl2.Text = $ktxt >> "%psfile%"
echo         $lbl2.BackColor = [System.Drawing.Color]::FromArgb(60,60,60) >> "%psfile%"
echo         $lbl2.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo         $lbl2.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo         $kw = if ($kcode -eq 'Back') { $(if ($script:kbNumpad) {100} else {120}) } else { $kbKeyW } >> "%psfile%"
echo         $lbl2.Size = New-Object System.Drawing.Size($kw, $kbKeyH) >> "%psfile%"
echo         $lbl2.Location = New-Object System.Drawing.Point($kbX, $kbStartY) >> "%psfile%"
echo         $lbl2.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter >> "%psfile%"
echo         $lbl2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle >> "%psfile%"
echo         $kbForm.Controls.Add($lbl2) >> "%psfile%"
echo         $kbLabels[$kcode] = $lbl2 >> "%psfile%"
echo         $kbX += $kw + $kbGap >> "%psfile%"
echo     } >> "%psfile%"
echo     $kbStartY += $kbKeyH + $kbGap >> "%psfile%"
echo } >> "%psfile%"

:: Dibujar numpad si aplica
echo if ($script:kbNumpad) { >> "%psfile%"
echo     $npX0 = 910; $npY = 55 >> "%psfile%"
echo     foreach ($nrow in $kbNumRows) { >> "%psfile%"
echo         $npX = $npX0 >> "%psfile%"
echo         foreach ($pair in $nrow) { >> "%psfile%"
echo             $ncode = $pair[0]; $ntxt = $pair[1] >> "%psfile%"
echo             $nLbl = New-Object System.Windows.Forms.Label >> "%psfile%"
echo             $nLbl.Text = $ntxt >> "%psfile%"
echo             $nLbl.BackColor = [System.Drawing.Color]::FromArgb(60,60,60) >> "%psfile%"
echo             $nLbl.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo             $nLbl.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo             $nkW = if ($ncode -eq 'NumPad0') {$kbKeyW*2+$kbGap} else {$kbKeyW} >> "%psfile%"
echo             $nLbl.Size = New-Object System.Drawing.Size($nkW, $kbKeyH) >> "%psfile%"
echo             $nLbl.Location = New-Object System.Drawing.Point($npX, $npY) >> "%psfile%"
echo             $nLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter >> "%psfile%"
echo             $nLbl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle >> "%psfile%"
echo             $kbForm.Controls.Add($nLbl) >> "%psfile%"
echo             $kbLabels[$ncode] = $nLbl >> "%psfile%"
echo             $npX += $nkW + $kbGap >> "%psfile%"
echo         } >> "%psfile%"
echo         $npY += $kbKeyH + $kbGap >> "%psfile%"
echo     } >> "%psfile%"
echo } >> "%psfile%"

echo $kbStatus = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $kbStatus.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $kbStatus.Font = New-Object System.Drawing.Font("Segoe UI", 11) >> "%psfile%"
echo $kbStatus.Size = New-Object System.Drawing.Size(1200, 135) >> "%psfile%"
echo $kbStatus.Location = New-Object System.Drawing.Point(40, 490) >> "%psfile%"
echo $kbForm.Controls.Add($kbStatus) >> "%psfile%"

echo function Get-KbDetectedNames { >> "%psfile%"
echo     $result = @() >> "%psfile%"
echo     for ($i=0; $i -lt $kbAllCodes.Count; $i++) { if ($kbPressed.Contains($kbAllCodes[$i])) { $result += $kbAllNames[$i] } } >> "%psfile%"
echo     return $result >> "%psfile%"
echo } >> "%psfile%"
echo function Get-KbMissingNames { >> "%psfile%"
echo     $result = @() >> "%psfile%"
echo     for ($i=0; $i -lt $kbAllCodes.Count; $i++) { if (-not $kbPressed.Contains($kbAllCodes[$i])) { $result += $kbAllNames[$i] } } >> "%psfile%"
echo     return $result >> "%psfile%"
echo } >> "%psfile%"
echo function Set-KbStatus([string]$instruction) { >> "%psfile%"
echo     $detected = @(Get-KbDetectedNames) >> "%psfile%"
echo     $missing = @(Get-KbMissingNames) >> "%psfile%"
echo     $detectedText = if ($detected.Count -gt 0) { $detected -join ', ' } else { 'Ninguna' } >> "%psfile%"
echo     $missingText = if ($missing.Count -gt 0) { $missing -join ', ' } else { 'Ninguna' } >> "%psfile%"
echo     $kbStatus.Text = "Detectadas: $detectedText`r`nNo detectadas: $missingText`r`n$instruction" >> "%psfile%"
echo } >> "%psfile%"
echo function Mark-KbTarget([string]$code) { >> "%psfile%"
echo     if (-not $kbLabels.ContainsKey($code)) { return } >> "%psfile%"
echo     $isNew = $kbPressed.Add($code) >> "%psfile%"
echo     if (-not $isNew) { return } >> "%psfile%"
echo     $kbLabels[$code].BackColor = [System.Drawing.Color]::FromArgb(0,190,0) >> "%psfile%"
echo     $kbLabels[$code].ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo     if ($script:kbConfirmPending) { >> "%psfile%"
echo         $script:kbConfirmPending = $false >> "%psfile%"
echo         $kbBtnDone.Text = 'REVISAR / FINALIZAR' >> "%psfile%"
echo         $kbBtnDone.BackColor = [System.Drawing.Color]::FromArgb(0,130,0) >> "%psfile%"
echo     } >> "%psfile%"
echo     Set-KbStatus 'Continue con las pendientes. Enter o boton para revisar.' >> "%psfile%"
echo } >> "%psfile%"
echo function Request-KbFinish { >> "%psfile%"
echo     $missing = @(Get-KbMissingNames) >> "%psfile%"
echo     if ($missing.Count -eq 0) { >> "%psfile%"
echo         $script:kbOutcome = 'Completa' >> "%psfile%"
echo         $script:kbMissingFinal = @() >> "%psfile%"
echo         $script:kbAllowClose = $true >> "%psfile%"
echo         $kbForm.Close() >> "%psfile%"
echo         return >> "%psfile%"
echo     } >> "%psfile%"
echo     if (-not $script:kbConfirmPending) { >> "%psfile%"
echo         $script:kbConfirmPending = $true >> "%psfile%"
echo         $kbBtnDone.Text = 'FINALIZAR CON PENDIENTES' >> "%psfile%"
echo         $kbBtnDone.BackColor = [System.Drawing.Color]::FromArgb(180,100,0) >> "%psfile%"
echo         Set-KbStatus 'Intente las teclas pendientes. Para finalizar con este resultado, presione Enter o el boton otra vez.' >> "%psfile%"
echo         return >> "%psfile%"
echo     } >> "%psfile%"
echo     $script:kbMissingFinal = @($missing) >> "%psfile%"
echo     $script:kbOutcome = 'Pendientes' >> "%psfile%"
echo     $script:kbAllowClose = $true >> "%psfile%"
echo     $kbForm.Close() >> "%psfile%"
echo } >> "%psfile%"

:: KeyDown: Enter controla revision; solo las teclas objetivo suman avance
echo $kbForm.Add_KeyDown({ >> "%psfile%"
echo     $_.SuppressKeyPress = $true >> "%psfile%"
echo     $_.Handled = $true >> "%psfile%"
echo     if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { Request-KbFinish; return } >> "%psfile%"
echo     $kk = $_.KeyCode.ToString() >> "%psfile%"
echo     if ($kk -eq 'Oemtilde' -or $kk -eq 'OemSemicolon') { $kk = 'Enye' } >> "%psfile%"
echo     Mark-KbTarget $kk >> "%psfile%"
echo }) >> "%psfile%"

echo $kbBtnDone = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $kbBtnDone.Text = 'REVISAR / FINALIZAR' >> "%psfile%"
echo $kbBtnDone.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $kbBtnDone.Size = New-Object System.Drawing.Size(380, 70) >> "%psfile%"
echo $kbBtnDone.Location = New-Object System.Drawing.Point(300, 405) >> "%psfile%"
echo $kbBtnDone.BackColor = [System.Drawing.Color]::FromArgb(0,130,0) >> "%psfile%"
echo $kbBtnDone.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $kbBtnDone.TabStop = $false >> "%psfile%"
echo $kbBtnDone.Add_MouseClick({ Request-KbFinish }) >> "%psfile%"
echo $kbForm.Controls.Add($kbBtnDone) >> "%psfile%"

echo $kbBtnCancel = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $kbBtnCancel.Text = 'CANCELAR PRUEBA' >> "%psfile%"
echo $kbBtnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $kbBtnCancel.Size = New-Object System.Drawing.Size(260, 70) >> "%psfile%"
echo $kbBtnCancel.Location = New-Object System.Drawing.Point(700, 405) >> "%psfile%"
echo $kbBtnCancel.BackColor = [System.Drawing.Color]::FromArgb(150,0,0) >> "%psfile%"
echo $kbBtnCancel.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $kbBtnCancel.TabStop = $false >> "%psfile%"
echo $kbBtnCancel.Add_MouseClick({ $script:kbOutcome = 'NoCompletada'; $script:kbAllowClose = $true; $kbForm.Close() }) >> "%psfile%"
echo $kbForm.Controls.Add($kbBtnCancel) >> "%psfile%"

echo $kbForm.Add_FormClosing({ >> "%psfile%"
echo     if (-not $script:kbAllowClose) { >> "%psfile%"
echo         $answer = [System.Windows.Forms.MessageBox]::Show('La prueba no esta completa. Desea cancelarla?', 'CityPC', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) >> "%psfile%"
echo         if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { >> "%psfile%"
echo             $script:kbOutcome = 'NoCompletada' >> "%psfile%"
echo             $script:kbAllowClose = $true >> "%psfile%"
echo         } else { >> "%psfile%"
echo             $_.Cancel = $true >> "%psfile%"
echo         } >> "%psfile%"
echo     } >> "%psfile%"
echo }) >> "%psfile%"

echo Set-KbStatus 'Presione cada tecla objetivo. Enter o boton para revisar.' >> "%psfile%"
echo $kbForm.ShowDialog() ^| Out-Null >> "%psfile%"

echo Log-Dual "   * Layout:       Espanol (con $kbEnyeName)" >> "%psfile%"
echo Log-Dual "   * Numpad:       $(if ($script:kbNumpad) {'Si'} else {'No'})" >> "%psfile%"
echo if ($script:kbOutcome -eq 'Completa') { >> "%psfile%"
echo     Log-Dual "   * Resultado:    TECLADO OK - Todas las teclas objetivo detectadas" "Green" >> "%psfile%"
echo } elseif ($script:kbOutcome -eq 'Pendientes') { >> "%psfile%"
echo     Log-Dual "   * Resultado:    No respondieron tras reintento: $($script:kbMissingFinal -join ', ')" "Yellow" >> "%psfile%"
echo } else { >> "%psfile%"
echo     Log-Dual "   * Resultado:    PRUEBA NO COMPLETADA" "Yellow" >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "" >> "%psfile%"
:: -------------------------------------------------------------------
:: SECCION 9: PRUEBA DE CAMARA (solo Laptop y AIO)
:: -------------------------------------------------------------------
echo } >> "%psfile%"
echo if ($tipoEquipo -ne 'CPU') { >> "%psfile%"
echo Log-Dual "=== 9. PRUEBA DE CAMARA ===" "Cyan" >> "%psfile%"

:: Detectar camaras instaladas via WMI (sin | en strings para no romper CMD)
echo $camaras = @(Get-CimInstance Win32_PnPEntity ^| Where-Object { $_.PNPClass -eq 'Camera' }) >> "%psfile%"
echo if ($camaras.Count -eq 0) { >> "%psfile%"
echo     $camaras = @(Get-CimInstance Win32_PnPEntity ^| Where-Object { $_.PNPClass -eq 'Image' -and $_.Status -eq 'OK' }) >> "%psfile%"
echo } >> "%psfile%"

echo if ($camaras.Count -gt 0) { >> "%psfile%"
echo     foreach ($cam in $camaras) { >> "%psfile%"
echo         Log-Dual "   * Detectada:     $($cam.Name)  [ $($cam.Status) ]" >> "%psfile%"
echo     } >> "%psfile%"
echo } else { >> "%psfile%"
echo     Log-Dual "   * Detectada:     No se encontro camara via PnP" "Yellow" >> "%psfile%"
echo } >> "%psfile%"

:: Intentar abrir la app de Camara de Windows
echo Write-Host "   Abriendo aplicacion de camara..." -ForegroundColor Cyan >> "%psfile%"
echo $camProc = Start-Process "microsoft.windows.camera:" -PassThru -ErrorAction SilentlyContinue >> "%psfile%"
echo Start-Sleep -Seconds 3 >> "%psfile%"

:: Ventana de pregunta al tecnico
echo $camForm = New-Object System.Windows.Forms.Form >> "%psfile%"
echo $camForm.FormBorderStyle = 'None' >> "%psfile%"
echo $camForm.Size = New-Object System.Drawing.Size(700, 320) >> "%psfile%"
echo $camForm.StartPosition = 'CenterScreen' >> "%psfile%"
echo $camForm.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 60) >> "%psfile%"
echo $camForm.TopMost = $true >> "%psfile%"

echo $camLbl = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $camLbl.Text = "PRUEBA DE CAMARA`n`nRevisa la app de Camara que se abrio.`nVes imagen de la camara correctamente?" >> "%psfile%"
echo $camLbl.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $camLbl.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $camLbl.Size = New-Object System.Drawing.Size(660, 180) >> "%psfile%"
echo $camLbl.Location = New-Object System.Drawing.Point(20, 20) >> "%psfile%"
echo $camLbl.TextAlign = 'MiddleCenter' >> "%psfile%"
echo $camForm.Controls.Add($camLbl) >> "%psfile%"

echo $camBtnOk = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $camBtnOk.Text = "CAMARA OK" >> "%psfile%"
echo $camBtnOk.Font = New-Object System.Drawing.Font("Segoe UI", 14) >> "%psfile%"
echo $camBtnOk.Size = New-Object System.Drawing.Size(200, 65) >> "%psfile%"
echo $camBtnOk.Location = New-Object System.Drawing.Point(80, 230) >> "%psfile%"
echo $camBtnOk.BackColor = [System.Drawing.Color]::FromArgb(0,140,0) >> "%psfile%"
echo $camBtnOk.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $camBtnOk.Add_Click({ $script:camRes = "OK"; $camForm.Close() }) >> "%psfile%"
echo $camForm.Controls.Add($camBtnOk) >> "%psfile%"

echo $camBtnNo = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $camBtnNo.Text = "SIN CAMARA / FALLA" >> "%psfile%"
echo $camBtnNo.Font = New-Object System.Drawing.Font("Segoe UI", 14) >> "%psfile%"
echo $camBtnNo.Size = New-Object System.Drawing.Size(240, 65) >> "%psfile%"
echo $camBtnNo.Location = New-Object System.Drawing.Point(360, 230) >> "%psfile%"
echo $camBtnNo.BackColor = [System.Drawing.Color]::FromArgb(160,0,0) >> "%psfile%"
echo $camBtnNo.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $camBtnNo.Add_Click({ $script:camRes = "CON FALLAS"; $camForm.Close() }) >> "%psfile%"
echo $camForm.Controls.Add($camBtnNo) >> "%psfile%"

echo $script:camRes = "Sin Probar" >> "%psfile%"
echo $camForm.ShowDialog() ^| Out-Null >> "%psfile%"

:: Cerrar app de camara
echo if ($camProc -and (-not $camProc.HasExited)) { >> "%psfile%"
echo     Stop-Process -Name "WindowsCamera" -Force -ErrorAction SilentlyContinue >> "%psfile%"
echo } >> "%psfile%"

echo Log-Dual "   * Resultado:    $script:camRes" >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: PIE DE REPORTE Y ENVIO
:: -------------------------------------------------------------------
echo } >> "%psfile%"
echo "==========================================" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "      FIN DEL REPORTE TECNICO" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "==========================================" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo Write-Host "`n[OK] REPORTE GENERADO EN ESCRITORIO" -ForegroundColor Green >> "%psfile%"

echo Write-Host "`n>>> ENVIANDO A LA NUBE (n8n)..." -ForegroundColor Cyan >> "%psfile%"

echo $cleanReport = [System.IO.File]::ReadAllText($archivo) >> "%psfile%"
echo $payload = @{ >> "%psfile%"
echo     ticket  = $ticket >> "%psfile%"
echo     tipo    = $tipoEquipo >> "%psfile%"
echo     fecha   = (Get-Date -Format "dd/MM/yyyy HH:mm") >> "%psfile%"
echo     tecnico = "" >> "%psfile%"
echo     reporte = $cleanReport >> "%psfile%"
echo } ^| ConvertTo-Json -Depth 5 >> "%psfile%"

echo $syncConfirmed = $false >> "%psfile%"
echo try { >> "%psfile%"
echo     [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 >> "%psfile%"
echo     $headers = @{ 'X-CityPC-USB-Token' = $env:CITYPC_USB_TOKEN } >> "%psfile%"
echo     $response = Invoke-RestMethod -Uri $webhookUrl -Method Post -Headers $headers -Body $payload -ContentType 'application/json' -TimeoutSec 60 -ErrorAction Stop >> "%psfile%"
echo     $reportHash = [string]$response.report_hash >> "%psfile%"
echo     $validResponse = ($null -ne $response) -and (-not ($response -is [System.Array])) -and ($response.ok -is [System.Boolean]) -and ($response.ok -eq $true) -and ([string]$response.code -ceq 'REPAIRSHOPR_COMMENT_CONFIRMED') -and ([string]$response.pipeline_code -ceq 'USB_DIAG_PIPELINE_ACCEPTED_V1') -and ([string]$response.ticket -ceq [string]$ticket) -and (-not [string]::IsNullOrWhiteSpace([string]$response.comment_id)) -and ($response.hidden -is [System.Boolean]) -and ($response.hidden -eq $true) -and ($response.citypc_persisted -is [System.Boolean]) -and ($response.citypc_persisted -eq $true) -and ($reportHash -cmatch '^[0-9a-fA-F]{64}$') -and ($response.citypc_accepted -is [System.Boolean]) -and ($response.citypc_accepted -eq $true) -and (-not [string]::IsNullOrWhiteSpace([string]$response.run_id)) >> "%psfile%"
echo     if (-not $validResponse) { throw 'n8n no confirmo el contrato completo del pipeline.' } >> "%psfile%"
echo     $syncConfirmed = $true >> "%psfile%"
echo     Write-Host '[OK] DATOS PERSISTIDOS Y PIPELINE CITYPC ACEPTADO.' -ForegroundColor Green >> "%psfile%"
echo } catch { >> "%psfile%"
echo     Write-Host '[X] REPORTE LOCAL CREADO, PERO EL SERVIDOR NO CONFIRMO LA SINCRONIZACION.' -ForegroundColor Red >> "%psfile%"
echo } >> "%psfile%"
echo if ($syncConfirmed) { exit 0 } >> "%psfile%"
echo exit 20 >> "%psfile%"

:: Ejecucion y limpieza
powershell -NoProfile -ExecutionPolicy Bypass -File "%psfile%"
set "DIAG_EXIT=!errorlevel!"
set "CITYPC_USB_TOKEN="
del "%psfile%"
del "%temp%\test_mic.wav" 2>nul
call :CLEAN_RUNTIME_FILES
if defined WIFI_CLEANUP_FAILED (
    color 0C
    echo.
    echo [X] NO SE PUDO RETIRAR EL PERFIL WI-FI TEMPORAL.
    echo Cierre esta ventana y avise a recepcion para retirar el perfil antes de entregar el equipo.
    if "!DIAG_EXIT!"=="0" set "DIAG_EXIT=21"
)

echo.
if "!DIAG_EXIT!"=="0" (
    echo ==========================================
    echo   PROCESO TERMINADO - NUBE CONFIRMADA
    echo ==========================================
) else (
    color 0C
    if "!DIAG_EXIT!"=="21" (
        echo ==========================================
        echo   NUBE CONFIRMADA - PERFIL WI-FI PENDIENTE
        echo ==========================================
        echo No entregue el equipo hasta retirar el perfil Wi-Fi temporal.
    ) else (
    echo ==========================================
    echo   REPORTE LOCAL CREADO - NUBE NO CONFIRMADA
    echo ==========================================
    echo Avise a recepcion. No vuelva a ejecutar para evitar duplicados.
    )
)
pause
exit /b !DIAG_EXIT!

:CLEAN_UPDATE_FILES
for %%F in ("%UPDATE_CHANNEL_PATH%" "%UPDATE_VALID_VERSION_PATH%" "%UPDATE_COMMIT_PATH%" "%UPDATE_MANIFEST_EXPECTED_HASH_PATH%" "%UPDATE_MANIFEST_PATH%" "%UPDATE_BAT_PATH%" "%UPDATE_HELPER_PATH%" "%UPDATE_COMMITTER_PATH%") do if exist "%%~F" del /F /Q "%%~F" >nul 2>&1
exit /b 0

:ACQUIRE_INSTANCE_LOCK
2>nul mkdir "%INSTANCE_LOCK%"
if not errorlevel 1 goto :INSTANCE_LOCK_ACQUIRED
set "INSTANCE_LOCK_OWNER="
if exist "%INSTANCE_LOCK%\owner.pid" set /p INSTANCE_LOCK_OWNER=<"%INSTANCE_LOCK%\owner.pid"
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:INSTANCE_LOCK_OWNER -notmatch '^[1-9][0-9]{0,9}$'){exit 0}; if($null -eq (Get-Process -Id ([int]$env:INSTANCE_LOCK_OWNER) -ErrorAction SilentlyContinue)){exit 0}; exit 1" >nul 2>&1
if errorlevel 1 exit /b 1
rmdir /S /Q "%INSTANCE_LOCK%" >nul 2>&1
2>nul mkdir "%INSTANCE_LOCK%"
if errorlevel 1 exit /b 1
:INSTANCE_LOCK_ACQUIRED
set "INSTANCE_LOCK_HELD=1"
call :CAPTURE_PARENT_PID
if errorlevel 1 (
    rmdir /S /Q "%INSTANCE_LOCK%" >nul 2>&1
    set "INSTANCE_LOCK_HELD="
    exit /b 1
)
>"%INSTANCE_LOCK%\owner.pid" echo %UPDATE_PARENT_PID%
if not exist "%INSTANCE_LOCK%\owner.pid" (
    rmdir /S /Q "%INSTANCE_LOCK%" >nul 2>&1
    set "INSTANCE_LOCK_HELD="
    exit /b 1
)
exit /b 0

:CAPTURE_PARENT_PID
set "UPDATE_PARENT_PID="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=Get-CimInstance Win32_Process -Filter ('ProcessId=' + $PID); if($null -eq $p){exit 1}; [Console]::Write($p.ParentProcessId)"`) do set "UPDATE_PARENT_PID=%%P"
powershell -NoProfile -ExecutionPolicy Bypass -Command "if($env:UPDATE_PARENT_PID -cmatch '^[1-9][0-9]{0,9}$'){exit 0}else{exit 1}" >nul 2>&1
exit /b %errorlevel%

:CLEAN_RUNTIME_FILES
if defined WIFI_READY (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%WIFI_HELPER%" -Action Cleanup >nul 2>&1
    if errorlevel 1 set "WIFI_CLEANUP_FAILED=1"
    set "WIFI_READY="
)
if defined WIFI_XML if exist "%WIFI_XML%" del /F /Q "%WIFI_XML%" >nul 2>&1
call :RELEASE_INSTANCE_LOCK
exit /b 0

:RELEASE_INSTANCE_LOCK
if defined INSTANCE_LOCK_HELD (
    2>nul rmdir /S /Q "%INSTANCE_LOCK%"
    set "INSTANCE_LOCK_HELD="
)
exit /b 0

:ERROR_CONFIG
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: Falta o es invalida la configuracion local junto al BAT.
echo Deben existir usbdiag.shared.local.cmd y wifi.local.cmd.
echo Copie los ejemplos, complete los valores privados y nunca los suba a Git.
pause
exit /b 30

:ERROR_WIFI_HELPER
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: Falta usbdiag-wifi-readiness.ps1 junto al BAT.
echo Recopie el paquete V8 completo. No continue con archivos parciales.
pause
exit /b 31

:ERROR_UPDATE_COMMITTER
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: Falta usbdiag-bundle-commit.ps1 junto al BAT.
echo Recopie el paquete V8 completo. No continue con archivos parciales.
pause
exit /b 32

:ERROR_UPDATE_HANDSHAKE
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: El reinicio de auto-update no presento token, version y hashes validos.
echo No se omitira la comprobacion ni se ejecutara un bundle sin confirmar.
pause
exit /b 35

:ERROR_UPDATE_CONFIRMATION
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: El bundle nuevo arranco, pero no pudo confirmar sus tres hashes.
echo La transaccion queda disponible para recuperacion en la siguiente apertura.
pause
exit /b 36

:ERROR_UPDATE_RECOVERY
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: Hay una actualizacion interrumpida y no se pudo iniciar su recuperacion.
echo No se ejecutara el diagnostico hasta restaurar o completar el bundle.
pause
exit /b 37

:ERROR_WIFI
set "WIFI_FAILED_RC=%WIFI_RC%"
call :CLEAN_RUNTIME_FILES
color 0C
echo.
echo ERROR: No se pudo dejar Wi-Fi listo de forma comprobable.
if "%WIFI_FAILED_RC%"=="41" echo Revise SSID y password en wifi.local.cmd.
if "%WIFI_FAILED_RC%"=="42" echo Revise el adaptador y el servicio WLAN de Windows.
if "%WIFI_FAILED_RC%"=="43" echo Windows no permitio crear el perfil temporal.
if "%WIFI_FAILED_RC%"=="44" echo Windows rechazo el perfil Wi-Fi; revise seguridad WPA2/AES.
if "%WIFI_FAILED_RC%"=="45" echo Windows rechazo la conexion; revise cobertura y credenciales.
if "%WIFI_FAILED_RC%"=="46" echo No se conecto al SSID exacto dentro de 45 segundos.
if "%WIFI_FAILED_RC%"=="47" echo El Wi-Fi no recibio una direccion IPv4 valida.
if "%WIFI_FAILED_RC%"=="48" echo Fallo DNS o HTTPS; revise internet, fecha y certificados.
echo No se ejecutara el diagnostico ni se reportara un falso OK.
pause
exit /b %WIFI_FAILED_RC%

:ERROR_WIFI_CLEANUP
call :RELEASE_INSTANCE_LOCK
color 0C
echo.
echo ERROR: No se pudo retirar el perfil Wi-Fi temporal del equipo.
echo No entregue el equipo hasta eliminar ese perfil y confirmar que ya no contiene la red CityPC.
pause
exit /b 49

:ERROR_ALREADY_RUNNING
color 0C
echo.
echo ERROR: Ya hay un Diagnostico CityPC abierto o quedo un bloqueo de una ejecucion interrumpida.
echo Cierre la otra ventana. Si no existe, elimine "%INSTANCE_LOCK%" y vuelva a intentar.
pause
exit /b 32

:ERROR_TICKET
color 0C
echo.
echo ERROR: El ticket debe tener 5 numeros.
pause
goto INICIO
