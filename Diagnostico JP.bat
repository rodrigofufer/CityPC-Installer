@echo off
cls
color 0B
setlocal enabledelayedexpansion

:: =========================================================
:: 1. CONFIGURACION GENERAL (EDITAR AQUI)
:: =========================================================
set "WEBHOOK_URL=https://n8n.srv1256135.hstgr.cloud/webhook/diagnostico-citypc"
set "WIFI_SSID=Citypc Slow"
set "WIFI_PASS=CitypcCitypc"
:: =========================================================

:: =========================================================
:: VERSION LOCAL Y AUTO-UPDATE
:: =========================================================
set "LOCAL_VER=1"
set "GITHUB_RAW=https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main"
set "REMOTE_FILE=Diagnostico_JP.bat"
set "VERSION_FILE=version_diagnostico.txt"

cls
echo.
echo ============================================================
echo      SISTEMA DE DIAGNOSTICO CITYPC (V%LOCAL_VER%)
echo ============================================================
echo.
echo Verificando si hay una version nueva...
echo.

:: Verificar internet
ping 8.8.8.8 -n 1 -w 2000 >nul 2>&1
if %errorlevel% neq 0 (
    echo [AVISO] Sin internet. Usando version actual V%LOCAL_VER%.
    echo.
    goto :skip_update_diag
)

:: Descargar version_diagnostico.txt de GitHub
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/%VERSION_FILE%','%temp%\citypc_ver_diag.txt') }catch{}" >nul 2>&1

if not exist "%temp%\citypc_ver_diag.txt" (
    echo [AVISO] No se pudo verificar version. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update_diag
)

:: Leer version remota y limpiar
set "REMOTE_VER="
set /p REMOTE_VER=<"%temp%\citypc_ver_diag.txt"
del /F /Q "%temp%\citypc_ver_diag.txt" >nul 2>&1

for /f "tokens=1 delims= " %%V in ("!REMOTE_VER!") do set "REMOTE_VER=%%V"

if not defined REMOTE_VER (
    echo [AVISO] No se pudo leer version remota. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update_diag
)

:: Comparar versiones
if "!REMOTE_VER!"=="!LOCAL_VER!" (
    echo [OK] Version V%LOCAL_VER% es la mas reciente.
    echo.
    goto :skip_update_diag
)

:: Hay version nueva - descargar
echo [!!] Nueva version disponible: V!REMOTE_VER! ^(actual: V%LOCAL_VER%^)
echo.
echo Descargando actualizacion...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/%REMOTE_FILE%','%temp%\Diagnostico_JP_update.bat') }catch{}" >nul 2>&1

if not exist "%temp%\Diagnostico_JP_update.bat" (
    echo [ERROR] No se pudo descargar. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update_diag
)

:: Reemplazar archivo actual
copy /Y "%temp%\Diagnostico_JP_update.bat" "%~dp0Diagnostico_JP.bat" >nul 2>&1
del /F /Q "%temp%\Diagnostico_JP_update.bat" >nul 2>&1

echo.
echo [OK] Actualizado a V!REMOTE_VER!. Reiniciando...
echo.
timeout /t 2 /nobreak >nul 2>&1

:: Reiniciar
start "" "%~dp0Diagnostico_JP.bat"
exit

:skip_update_diag


:INICIO
cls
echo ==========================================
echo      SISTEMA DE DIAGNOSTICO CITYPC (V%LOCAL_VER%)
echo ==========================================
echo.
set "TICKET="
set /p "TICKET=Ingrese # Ticket (5 digitos): "

if "%TICKET:~4,1%"=="" goto ERROR_TICKET
if not "%TICKET:~5,1%"=="" goto ERROR_TICKET

:: ---------------------------------------------------------
:: 2. CONEXION WI-FI
:: ---------------------------------------------------------
echo.
echo [1/8] Conectando a Wi-Fi "%WIFI_SSID%"...
set "WIFI_XML=%temp%\wifi_config.xml"
if exist "%WIFI_XML%" del "%WIFI_XML%"
(
echo ^<?xml version="1.0"?^>
echo ^<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1"^>
echo   ^<name^>%WIFI_SSID%^</name^>
echo   ^<SSIDConfig^>
echo     ^<SSID^>
echo       ^<name^>%WIFI_SSID%^</name^>
echo     ^</SSID^>
echo   ^</SSIDConfig^>
echo   ^<connectionType^>ESS^</connectionType^>
echo   ^<connectionMode^>auto^</connectionMode^>
echo   ^<MSM^>
echo     ^<security^>
echo       ^<authEncryption^>
echo         ^<authentication^>WPA2PSK^</authentication^>
echo         ^<encryption^>AES^</encryption^>
echo         ^<useOneX^>false^</useOneX^>
echo       ^</authEncryption^>
echo       ^<sharedKey^>
echo         ^<keyType^>passPhrase^</keyType^>
echo         ^<protected^>false^</protected^>
echo         ^<keyMaterial^>%WIFI_PASS%^</keyMaterial^>
echo       ^</sharedKey^>
echo     ^</security^>
echo   ^</MSM^>
echo ^</WLANProfile^>
) > "%WIFI_XML%"
netsh wlan add profile filename="%WIFI_XML%" >nul
netsh wlan connect name="%WIFI_SSID%" >nul
del "%WIFI_XML%"

echo.
echo [2/8] Preparando herramientas...
set "psfile=%temp%\diag_v5.ps1"
if exist "%psfile%" del "%psfile%"

:: ---------------------------------------------------------
:: 3. GENERANDO SCRIPT POWERSHELL
:: ---------------------------------------------------------
echo $ErrorActionPreference = 'SilentlyContinue' >> "%psfile%"
echo [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 >> "%psfile%"
echo $ticket = "%TICKET%" >> "%psfile%"
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
echo $mem   = Get-CimInstance Win32_PhysicalMemory >> "%psfile%"
echo $slots = Get-CimInstance Win32_PhysicalMemoryArray >> "%psfile%"
echo Log-Dual "=== 3. MEMORIA RAM ===" "Cyan" >> "%psfile%"
echo $totalMem   = [math]::Round(($mem ^| Measure-Object -Property Capacity -Sum).Sum / 1GB) >> "%psfile%"
echo $usados     = @($mem).Count >> "%psfile%"
echo $libresSlot = $slots.MemoryDevices - $usados >> "%psfile%"
echo Log-Dual "   * Instalada:     $totalMem GB  ($($mem[0].Speed) MHz)" >> "%psfile%"
echo Log-Dual "   * Ranuras:       $usados Ocupadas / $libresSlot Libres" >> "%psfile%"
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
:: SECCION 6: DIAGNOSTICO DE PANTALLA
:: -------------------------------------------------------------------
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
:: SECCION 8: PRUEBA DE TECLADO (NUEVA - forma visual interactiva)
:: -------------------------------------------------------------------
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
:: Fila numeros + BKSP
echo $rowNums = @(@('D1','1'),@('D2','2'),@('D3','3'),@('D4','4'),@('D5','5'),@('D6','6'),@('D7','7'),@('D8','8'),@('D9','9'),@('D0','0'),@('Back','BKSP')) >> "%psfile%"
:: Fila QWERTY
echo $rowQ = @(@('Q','Q'),@('W','W'),@('E','E'),@('R','R'),@('T','T'),@('Y','Y'),@('U','U'),@('I','I'),@('O','O'),@('P','P')) >> "%psfile%"
:: Fila ASDF con Enye
echo $rowA = @(@('A','A'),@('S','S'),@('D','D'),@('F','F'),@('G','G'),@('H','H'),@('J','J'),@('K','K'),@('L','L'),@('OemSemicolon','ENY')) >> "%psfile%"
:: Fila ZXCV
echo $rowZ = @(@('Z','Z'),@('X','X'),@('C','C'),@('V','V'),@('B','B'),@('N','N'),@('M','M')) >> "%psfile%"

echo $kbMainRows = @($rowNums, $rowQ, $rowA, $rowZ) >> "%psfile%"

:: Numpad rows
echo $rowNP1 = @(@('NumPad7','N7'),@('NumPad8','N8'),@('NumPad9','N9')) >> "%psfile%"
echo $rowNP2 = @(@('NumPad4','N4'),@('NumPad5','N5'),@('NumPad6','N6')) >> "%psfile%"
echo $rowNP3 = @(@('NumPad1','N1'),@('NumPad2','N2'),@('NumPad3','N3')) >> "%psfile%"
echo $rowNP4 = @(,@('NumPad0','N0')) >> "%psfile%"
echo $kbNumRows = @($rowNP1,$rowNP2,$rowNP3,$rowNP4) >> "%psfile%"

:: Lista de todos los codigos para el reporte final
echo $kbAllCodes = @('D1','D2','D3','D4','D5','D6','D7','D8','D9','D0','Back','Q','W','E','R','T','Y','U','I','O','P','A','S','D','F','G','H','J','K','L','OemSemicolon','Z','X','C','V','B','N','M') >> "%psfile%"
echo $kbAllNames = @('1','2','3','4','5','6','7','8','9','0','BKSP','Q','W','E','R','T','Y','U','I','O','P','A','S','D','F','G','H','J','K','L','ENY','Z','X','C','V','B','N','M') >> "%psfile%"
echo if ($script:kbNumpad) { >> "%psfile%"
echo     $kbAllCodes += @('NumPad7','NumPad8','NumPad9','NumPad4','NumPad5','NumPad6','NumPad1','NumPad2','NumPad3','NumPad0') >> "%psfile%"
echo     $kbAllNames += @('N7','N8','N9','N4','N5','N6','N1','N2','N3','N0') >> "%psfile%"
echo } >> "%psfile%"

:: ------ VENTANA PRINCIPAL ------
echo $kbForm = New-Object System.Windows.Forms.Form >> "%psfile%"
echo $kbForm.FormBorderStyle = 'None' >> "%psfile%"
echo $kbForm.WindowState = 'Maximized' >> "%psfile%"
echo $kbForm.BackColor = [System.Drawing.Color]::FromArgb(18,18,18) >> "%psfile%"
echo $kbForm.TopMost = $true >> "%psfile%"
echo $kbForm.KeyPreview = $true >> "%psfile%"
echo $kbForm.AcceptButton = $null >> "%psfile%"
echo $kbForm.CancelButton = $null >> "%psfile%"

echo $npAviso = if ($script:kbNumpad) {'  NUMPAD: activa NumLock antes de probar'} else {''} >> "%psfile%"
echo $kbTitle = New-Object System.Windows.Forms.Label >> "%psfile%"
echo $kbTitle.Text = "PRUEBA DE TECLADO - ESPANOL  --  Presiona cada tecla  --  Solo click FINALIZAR para salir$npAviso" >> "%psfile%"
echo $kbTitle.ForeColor = [System.Drawing.Color]::Cyan >> "%psfile%"
echo $kbTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $kbTitle.AutoSize = $true >> "%psfile%"
echo $kbTitle.Location = New-Object System.Drawing.Point(20, 12) >> "%psfile%"
echo $kbForm.Controls.Add($kbTitle) >> "%psfile%"

echo $kbLabels  = @{} >> "%psfile%"
echo $kbPressed = [System.Collections.Generic.HashSet[string]]::new() >> "%psfile%"
echo $kbKeyW = 82; $kbKeyH = 82; $kbGap = 6; $kbStartY = 60 >> "%psfile%"

:: Dibujar teclas principales
echo foreach ($row in $kbMainRows) { >> "%psfile%"
echo     $kbX = 40 >> "%psfile%"
echo     foreach ($pair in $row) { >> "%psfile%"
echo         $kcode = $pair[0]; $ktxt = $pair[1] >> "%psfile%"
echo         $lbl2 = New-Object System.Windows.Forms.Label >> "%psfile%"
echo         $lbl2.Text = $ktxt >> "%psfile%"
echo         $lbl2.BackColor = [System.Drawing.Color]::FromArgb(60,60,60) >> "%psfile%"
echo         $lbl2.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo         $lbl2.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo         $kw = if ($kcode -eq 'Back') {140} else {$kbKeyW} >> "%psfile%"
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
echo     $npX0 = 1060; $npY = 60 >> "%psfile%"
echo     foreach ($nrow in $kbNumRows) { >> "%psfile%"
echo         $npX = $npX0 >> "%psfile%"
echo         foreach ($pair in $nrow) { >> "%psfile%"
echo             $ncode = $pair[0]; $ntxt = $pair[1] >> "%psfile%"
echo             $nLbl = New-Object System.Windows.Forms.Label >> "%psfile%"
echo             $nLbl.Text = $ntxt >> "%psfile%"
echo             $nLbl.BackColor = [System.Drawing.Color]::FromArgb(60,60,60) >> "%psfile%"
echo             $nLbl.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo             $nLbl.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
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

:: KeyDown: bloquear TODO excepto las teclas de prueba. Insert = NumPad0 con NumLock OFF
echo $kbForm.Add_KeyDown({ >> "%psfile%"
echo     $_.SuppressKeyPress = $true >> "%psfile%"
echo     $kk = $_.KeyCode.ToString() >> "%psfile%"
echo     if ($kk -eq 'Insert' -and $kbLabels.ContainsKey('NumPad0')) { $kk = 'NumPad0' } >> "%psfile%"
echo     if ($kk -eq 'Oemtilde' -or $kk -eq 'OemSemicolon') { $kk = 'OemSemicolon' } >> "%psfile%"
echo     if ($kbLabels.ContainsKey($kk)) { >> "%psfile%"
echo         $kbLabels[$kk].BackColor = [System.Drawing.Color]::FromArgb(0,190,0) >> "%psfile%"
echo         $kbLabels[$kk].ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo         $kbPressed.Add($kk) ^| Out-Null >> "%psfile%"
echo     } >> "%psfile%"
echo }) >> "%psfile%"

echo $kbBtnDone = New-Object System.Windows.Forms.Button >> "%psfile%"
echo $kbBtnDone.Text = "FINALIZAR PRUEBA DE TECLADO" >> "%psfile%"
echo $kbBtnDone.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold) >> "%psfile%"
echo $kbBtnDone.Size = New-Object System.Drawing.Size(420, 75) >> "%psfile%"
echo $kbBtnDone.Location = New-Object System.Drawing.Point(430, 510) >> "%psfile%"
echo $kbBtnDone.BackColor = [System.Drawing.Color]::FromArgb(0,130,0) >> "%psfile%"
echo $kbBtnDone.ForeColor = [System.Drawing.Color]::White >> "%psfile%"
echo $kbBtnDone.Add_Click({ $kbForm.Close() }) >> "%psfile%"
echo $kbForm.Controls.Add($kbBtnDone) >> "%psfile%"

echo $kbForm.ShowDialog() ^| Out-Null >> "%psfile%"

:: Reporte de teclas sin probar
echo $kbMissing = @() >> "%psfile%"
echo for ($i=0; $i -lt $kbAllCodes.Count; $i++) { >> "%psfile%"
echo     if (-not $kbPressed.Contains($kbAllCodes[$i])) { $kbMissing += $kbAllNames[$i] } >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "   * Layout:       Espanol (con Enye)" >> "%psfile%"
echo Log-Dual "   * Numpad:       $(if ($script:kbNumpad) {'Si'} else {'No'})" >> "%psfile%"
echo if ($kbMissing.Count -eq 0) { >> "%psfile%"
echo     Log-Dual "   * Resultado:    TECLADO OK - Todas las teclas probadas" "Green" >> "%psfile%"
echo } else { >> "%psfile%"
echo     Log-Dual "   * Sin probar:   $($kbMissing -join ', ')" "Yellow" >> "%psfile%"
echo } >> "%psfile%"
echo Log-Dual "" >> "%psfile%"
:: -------------------------------------------------------------------
:: SECCION 9: PRUEBA DE CAMARA (NUEVA)
:: -------------------------------------------------------------------
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
echo if ($camProc -and !$camProc.HasExited) { >> "%psfile%"
echo     Stop-Process -Name "WindowsCamera" -Force -ErrorAction SilentlyContinue >> "%psfile%"
echo } >> "%psfile%"

echo Log-Dual "   * Resultado:    $script:camRes" >> "%psfile%"
echo Log-Dual "" >> "%psfile%"

:: -------------------------------------------------------------------
:: PIE DE REPORTE Y ENVIO
:: -------------------------------------------------------------------
echo "==========================================" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "      FIN DEL REPORTE TECNICO" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo "==========================================" ^| Out-File -FilePath $archivo -Append -Encoding UTF8 >> "%psfile%"
echo Write-Host "`n[OK] REPORTE GENERADO EN ESCRITORIO" -ForegroundColor Green >> "%psfile%"

echo Write-Host "`n>>> ENVIANDO A LA NUBE (n8n)..." -ForegroundColor Cyan >> "%psfile%"

echo $cleanReport = [System.IO.File]::ReadAllText($archivo) >> "%psfile%"
echo $payload = @{ >> "%psfile%"
echo     ticket  = $ticket >> "%psfile%"
echo     fecha   = (Get-Date -Format "dd/MM/yyyy HH:mm") >> "%psfile%"
echo     tecnico = "" >> "%psfile%"
echo     reporte = $cleanReport >> "%psfile%"
echo } ^| ConvertTo-Json -Depth 5 >> "%psfile%"

echo if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { >> "%psfile%"
echo     try { >> "%psfile%"
echo         [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 >> "%psfile%"
echo         $response = Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $payload -ContentType "application/json" -TimeoutSec 15 >> "%psfile%"
echo         Write-Host "[OK] DATOS SINCRONIZADOS CORRECTAMENTE." -ForegroundColor Green >> "%psfile%"
echo     } catch { >> "%psfile%"
echo         if ($_.Exception.Message -like "*time*") { >> "%psfile%"
echo             Write-Host "[!]  Enviado, pero sin confirmacion rapida del servidor." -ForegroundColor Yellow >> "%psfile%"
echo         } else { >> "%psfile%"
echo             Write-Host "[X] ERROR: " + $_.Exception.Message -ForegroundColor Red >> "%psfile%"
echo         } >> "%psfile%"
echo     } >> "%psfile%"
echo } else { >> "%psfile%"
echo     Write-Host "[X] SIN INTERNET. Reporte guardado solo localmente." -ForegroundColor Red >> "%psfile%"
echo } >> "%psfile%"

:: Ejecucion y limpieza
powershell -NoProfile -ExecutionPolicy Bypass -File "%psfile%"
del "%psfile%"
del "%temp%\test_mic.wav" 2>nul

echo.
echo ==========================================
echo   PROCESO TERMINADO - PUEDE CERRAR
echo ==========================================
pause
exit

:ERROR_TICKET
color 0C
echo.
echo ERROR: El ticket debe tener 5 numeros.
pause
goto INICIO
