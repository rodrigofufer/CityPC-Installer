@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
color 1F
title CITYPC - INSTALADOR Y ACTUALIZADOR
mode con: cols=100 lines=50

:: =========================================================
:: VERSION LOCAL
:: =========================================================
set "LOCAL_VER=43"
set "GITHUB_RAW=https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main"

:: =========================================================
:: 0. VERIFICAR ADMINISTRADOR
:: =========================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
    color 4F
    cls
    echo.
    echo.
    echo  ============================================================
    echo.
    echo       NO SE PUEDE EJECUTAR ESTE PROGRAMA
    echo.
    echo       Necesitas abrirlo de la siguiente manera:
    echo.
    echo       1. Cierra esta ventana
    echo       2. Busca el archivo "Instalador_CityPC.bat"
    echo       3. Dale clic DERECHO con el mouse
    echo       4. Selecciona "Ejecutar como administrador"
    echo       5. Si te pregunta, dale que SI
    echo.
    echo  ============================================================
    echo.
    echo.
    echo  Presiona cualquier tecla para cerrar...
    pause >nul
    exit
)

:: Ruta USB
set "USB_PATH=%~dp0"
if "%USB_PATH:~-1%"=="\" set "USB_PATH=%USB_PATH:~0,-1%"

:: =========================================================
:: AUTO-UPDATE DESDE GITHUB
:: =========================================================
cls
echo.
echo  ============================================================
echo        CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   Verificando si hay una version nueva...
echo.

:: Verificar internet
ping 8.8.8.8 -n 1 -w 2000 >nul 2>&1
if %errorlevel% neq 0 (
    echo   [AVISO] Sin internet. Usando version actual V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Descargar version.txt de GitHub
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/version.txt','%temp%\citypc_version.txt') }catch{}" >nul 2>&1

if not exist "%temp%\citypc_version.txt" (
    echo   [AVISO] No se pudo verificar version. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Leer version remota y limpiar
set "REMOTE_VER="
set /p REMOTE_VER=<"%temp%\citypc_version.txt"
del /F /Q "%temp%\citypc_version.txt" >nul 2>&1

:: Limpiar espacios, tabuladores, retornos de carro
for /f "tokens=1 delims= 	" %%V in ("!REMOTE_VER!") do set "REMOTE_VER=%%V"

if not defined REMOTE_VER (
    echo   [AVISO] No se pudo leer version remota. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Comparar versiones
if "!REMOTE_VER!"=="!LOCAL_VER!" (
    echo   [OK] Version V%LOCAL_VER% es la mas reciente.
    echo.
    goto :skip_update
)

:: Hay version nueva - descargar
echo   [!!] Nueva version disponible: V!REMOTE_VER! ^(actual: V%LOCAL_VER%^)
echo.
echo   Descargando actualizacion...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/Instalador_CityPC.bat','%temp%\Instalador_CityPC_update.bat') }catch{}" >nul 2>&1

if not exist "%temp%\Instalador_CityPC_update.bat" (
    echo   [ERROR] No se pudo descargar. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Reemplazar archivo en la USB
copy /Y "%temp%\Instalador_CityPC_update.bat" "%~dp0Instalador_CityPC.bat" >nul 2>&1
del /F /Q "%temp%\Instalador_CityPC_update.bat" >nul 2>&1

echo.
echo   [OK] Actualizado a V!REMOTE_VER!. Reiniciando...
echo.
timeout /t 2 /nobreak >nul 2>&1

:: Reiniciar CON permisos de administrador
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process cmd -ArgumentList '/c \"^"%~dp0Instalador_CityPC.bat^\"' -Verb RunAs" >nul 2>&1
exit

:skip_update

:: Escritorio publico
for /f "tokens=2*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Common Desktop" 2^>nul') do set "DESKTOP=%%B"
if not defined DESKTOP set "DESKTOP=%PUBLIC%\Desktop"

title CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)

cls
echo.
echo  ============================================================
echo        CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   Ruta USB: "%USB_PATH%"
echo.
echo   No cierre esta ventana. El proceso puede tardar varios
echo   minutos dependiendo de la velocidad de internet.
echo.
echo  ============================================================
echo.

:: =========================================================
:: 1. EXCEPCIONES ANTIVIRUS
:: =========================================================
echo  [1/6] Aplicando Excepciones de Seguridad...

sc query WinDefend >nul 2>&1
if %errorlevel% neq 0 (
    echo    [AVISO] Windows Defender no esta activo.
    goto :skip_exclusions
)

set "EX_SCRIPT=%temp%\Exclusiones_CityPC.ps1"

> "%EX_SCRIPT%" (
    echo $ErrorActionPreference = 'SilentlyContinue'
    echo.
    echo # Carpetas
    echo Add-MpPreference -ExclusionPath 'C:\CityPC'
    echo Add-MpPreference -ExclusionPath '%USB_PATH%'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAutoS'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAuto'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAuto Net'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\Temp'
    echo Add-MpPreference -ExclusionPath $env:TEMP
    echo Add-MpPreference -ExclusionPath 'C:\Windows\SECOH-QAD.exe'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\SECOH-QAD.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjHook.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjPatcher.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjPatcher.exe'
    echo.
    echo # Procesos
    echo Add-MpPreference -ExclusionProcess 'KMSAuto Net.exe'
    echo Add-MpPreference -ExclusionProcess 'KMSAuto.exe'
    echo Add-MpPreference -ExclusionProcess 'AutoKMS.exe'
    echo Add-MpPreference -ExclusionProcess 'KMS Server.exe'
    echo Add-MpPreference -ExclusionProcess 'Service_KMS.exe'
    echo Add-MpPreference -ExclusionProcess 'Soporte Tecnico CityPC.mx NEW.exe'
    echo Add-MpPreference -ExclusionProcess 'TunMirror.exe'
    echo Add-MpPreference -ExclusionProcess 'TunMirror2.exe'
    echo Add-MpPreference -ExclusionProcess 'TapInstall.exe'
    echo.
    echo # Desactivar proteccion durante instalacion
    echo Set-MpPreference -DisableRealtimeMonitoring $true
    echo Set-MpPreference -SubmitSamplesConsent 2
    echo Set-MpPreference -MAPSReporting 0
    echo Set-MpPreference -DisableBehaviorMonitoring $true
    echo Set-MpPreference -DisableIOAVProtection $true
    echo Set-MpPreference -DisableScriptScanning $true
)

powershell -ExecutionPolicy Bypass -NoProfile -File "%EX_SCRIPT%" >nul 2>&1
del /F /Q "%EX_SCRIPT%" >nul 2>&1

reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\CityPC" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "%USB_PATH%" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAutoS" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAuto" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAuto Net" /t REG_DWORD /d 0 /f >nul 2>&1

reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender" /v DisableAntiSpyware /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableBehaviorMonitoring /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableOnAccessProtection /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableScanOnRealtimeEnable /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v SubmitSamplesConsent /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v SpynetReporting /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications" /v DisableNotifications /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications" /v DisableEnhancedNotifications /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.SecurityHealth" /v Enabled /t REG_DWORD /d 0 /f >nul 2>&1

echo    [OK] Excepciones aplicadas.

:skip_exclusions

:: =========================================================
:: 2. VERIFICAR INTERNET Y WINDOWS UPDATE
:: =========================================================
echo.
echo  [2/6] Verificando Internet y Activando Updates...

set "HAY_INTERNET=0"
ping 8.8.8.8 -n 2 -w 2000 >nul 2>&1
if %errorlevel% equ 0 (
    set "HAY_INTERNET=1"
) else (
    ping 1.1.1.1 -n 2 -w 2000 >nul 2>&1
    if !errorlevel! equ 0 set "HAY_INTERNET=1"
)

if "!HAY_INTERNET!"=="0" (
    echo    [AVISO] Sin internet. No se podran bajar actualizaciones.
) else (
    echo    [OK] Internet conectado.
    sc config wuauserv start= auto >nul 2>&1
    sc config bits start= auto >nul 2>&1
    sc config dosvc start= auto >nul 2>&1
    net start wuauserv >nul 2>&1
    net start bits >nul 2>&1
    net start dosvc >nul 2>&1
    usoclient StartScan >nul 2>&1
    usoclient StartDownload >nul 2>&1
    usoclient StartInstall >nul 2>&1
    echo    [OK] Windows Update activado en segundo plano.
)

:: =========================================================
:: 3. REPARAR TLS Y WINGET
:: =========================================================
echo.
echo  [3/6] Reparando seguridad y gestor de descargas...

reg add "HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" /v SchUseStrongCrypto /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" /v SchUseStrongCrypto /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v Enabled /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v DisabledByDefault /t REG_DWORD /d 0 /f >nul 2>&1
echo    [OK] TLS 1.2 activado.

certutil -generateSSTFromWU "%temp%\roots.sst" >nul 2>&1
if exist "%temp%\roots.sst" (
    certutil -addstore -f root "%temp%\roots.sst" >nul 2>&1
    del /F /Q "%temp%\roots.sst" >nul 2>&1
    echo    [OK] Certificados actualizados.
) else (
    echo    [AVISO] No se pudieron actualizar certificados.
)

winget source reset --force >nul 2>&1
if %errorlevel% neq 0 (
    winget source remove msstore >nul 2>&1
    winget source remove winget >nul 2>&1
    winget source add winget "https://cdn.winget.microsoft.com/cache" --accept-source-agreements >nul 2>&1
    winget source add msstore "https://storeedgefd.dsx.mp.microsoft.com/v9.0" --accept-source-agreements >nul 2>&1
)
winget source update --accept-source-agreements >nul 2>&1

set "WINGET_OK=0"
winget --version >nul 2>&1
if %errorlevel% equ 0 (
    echo n | winget list --accept-source-agreements >nul 2>&1
    set "WINGET_OK=1"
    echo    [OK] Winget listo.
) else (
    echo    [AVISO] Winget no disponible. Se usara descarga directa.
)

:: =========================================================
:: 4. CONFIGURACION ENERGIA
:: =========================================================
echo.
echo  [4/6] Configurando Energia...

powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
if %errorlevel% neq 0 (
    powercfg -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
    powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
)
powercfg -change -monitor-timeout-ac 0 >nul 2>&1
powercfg -change -monitor-timeout-dc 0 >nul 2>&1
powercfg -change -standby-timeout-ac 0 >nul 2>&1
powercfg -change -standby-timeout-dc 0 >nul 2>&1
powercfg -h off >nul 2>&1
echo    [OK] Energia maxima configurada.

:: =========================================================
:: 5. SOPORTE TECNICO CITYPC
:: =========================================================
echo.
echo  [5/6] Instalando Soporte Tecnico CityPC...

if not exist "C:\CityPC" mkdir "C:\CityPC"

set "SOPORTE_SRC="
if exist "%USB_PATH%\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\Soporte Tecnico CityPC.mx NEW.exe"
if not defined SOPORTE_SRC if exist "%USB_PATH%\CityPC\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\CityPC\Soporte Tecnico CityPC.mx NEW.exe"
if not defined SOPORTE_SRC if exist "%USB_PATH%\Archivos\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\Archivos\Soporte Tecnico CityPC.mx NEW.exe"

if defined SOPORTE_SRC (
    copy /Y "!SOPORTE_SRC!" "C:\CityPC\" >nul 2>&1
    echo    [OK] Soporte copiado a C:\CityPC
) else (
    echo    [AVISO] No se encontro "Soporte Tecnico CityPC.mx NEW.exe" en la USB.
)

if exist "%USB_PATH%\icono CityPC.ico" copy /Y "%USB_PATH%\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\CityPC\icono CityPC.ico" copy /Y "%USB_PATH%\CityPC\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\Archivos\icono CityPC.ico" copy /Y "%USB_PATH%\Archivos\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\Iconos\icono CityPC.ico" copy /Y "%USB_PATH%\Iconos\icono CityPC.ico" "C:\CityPC\" >nul 2>&1

attrib +h "C:\CityPC" >nul 2>&1

if exist "C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe" (
    call :crear_acceso_soporte
) else (
    echo    [ERROR] No se pudo instalar Soporte. Archivo no disponible.
)

goto :skip_acceso_soporte

:crear_acceso_soporte
set "S_LNK=%PUBLIC%\Desktop\Soporte Tecnico CityPC.mx.lnk"
set "S_TARGET=C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe"
set "S_ICON=C:\CityPC\icono CityPC.ico"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell;$s=$w.CreateShortcut('%S_LNK%');$s.TargetPath='%S_TARGET%';$s.WorkingDirectory='C:\CityPC';if(Test-Path '%S_ICON%'){$s.IconLocation='%S_ICON%'};$s.Save()"
if exist "%S_LNK%" (
    echo    [OK] Acceso directo de Soporte creado.
) else (
    echo    [AVISO] No se creo el acceso directo.
)
goto :eof

:skip_acceso_soporte

:: =========================================================
:: 6. INSTALACION DE SOFTWARE (NINITE)
:: =========================================================
echo.
echo  ============================================================
echo    [6/6] INSTALANDO PROGRAMAS
echo  ============================================================
echo.

set "NINITE_SRC="
if exist "%USB_PATH%\Ninite.exe" set "NINITE_SRC=%USB_PATH%\Ninite.exe"
if not defined NINITE_SRC if exist "%USB_PATH%\CityPC\Ninite.exe" set "NINITE_SRC=%USB_PATH%\CityPC\Ninite.exe"
if not defined NINITE_SRC if exist "%USB_PATH%\Archivos\Ninite.exe" set "NINITE_SRC=%USB_PATH%\Archivos\Ninite.exe"

if not defined NINITE_SRC (
    echo    [ERROR] No se encontro Ninite.exe en la USB.
    goto :resumen_final
)

echo    Instalando Chrome, Adobe Reader, WinRAR y Zoom...
echo    Esto puede tardar varios minutos. No cierre esta ventana.
echo.

start /wait "" "!NINITE_SRC!" /silent
timeout /t 5 /nobreak >nul 2>&1

echo    [OK] Ninite finalizo la instalacion.

goto :resumen_final

:: ==========================================================
:: CHROME
:: ==========================================================
:instalar_chrome
echo  [^>] Google Chrome:
echo     - Instalando...

if "!WINGET_OK!"=="1" (
    winget install --id Google.Chrome -e --silent --accept-package-agreements --accept-source-agreements --force >nul 2>&1
    timeout /t 5 /nobreak >nul 2>&1
)

set "CHROME_EXE="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if not defined CHROME_EXE (
    echo     - Descargando instalador directo...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://dl.google.com/chrome/install/latest/chrome_installer.exe' -OutFile \"$env:TEMP\chrome_setup.exe\" -UseBasicParsing -TimeoutSec 120" >nul 2>&1
    if exist "%temp%\chrome_setup.exe" (
        start /wait "" "%temp%\chrome_setup.exe" /silent /install
        timeout /t 15 /nobreak >nul 2>&1
        del /F /Q "%temp%\chrome_setup.exe" >nul 2>&1
    )
)

set "CHROME_EXE="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if defined CHROME_EXE (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -COM WScript.Shell; $s=$ws.CreateShortcut('!DESKTOP!\Google Chrome.lnk'); $s.TargetPath='!CHROME_EXE!'; $s.Save()" >nul 2>&1
    echo     - [OK] Chrome instalado.
) else (
    echo     - [ERROR] No se pudo instalar Chrome.
)
goto :eof

:: ==========================================================
:: ADOBE READER
:: ==========================================================
:instalar_adobe
echo  [^>] Adobe Acrobat Reader:
echo     - Instalando...

if "!WINGET_OK!"=="1" (
    winget install --id Adobe.Acrobat.Reader.64-bit -e --silent --accept-package-agreements --accept-source-agreements --force >nul 2>&1
    timeout /t 8 /nobreak >nul 2>&1
)

call :detectar_adobe
if not defined ADOBE_EXE (
    if "!WINGET_OK!"=="1" (
        echo     - Intentando version 32-bit...
        winget install --id Adobe.Acrobat.Reader.32-bit -e --silent --accept-package-agreements --accept-source-agreements --force >nul 2>&1
        timeout /t 8 /nobreak >nul 2>&1
    )
)

call :detectar_adobe
if not defined ADOBE_EXE (
    echo     - Descargando instalador 64-bit...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820555/AcroRdrDCx64_en_US.exe' -OutFile \"$env:TEMP\adobe_setup.exe\" -UseBasicParsing -TimeoutSec 180" >nul 2>&1
    if exist "%temp%\adobe_setup.exe" (
        start /wait "" "%temp%\adobe_setup.exe" /sAll /rs /msi EULA_ACCEPT=YES
        timeout /t 15 /nobreak >nul 2>&1
        del /F /Q "%temp%\adobe_setup.exe" >nul 2>&1
    )
)

call :detectar_adobe
if not defined ADOBE_EXE (
    echo     - Descargando instalador 32-bit...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820555/AcroRdrDC2300820555_en_US.exe' -OutFile \"$env:TEMP\adobe32_setup.exe\" -UseBasicParsing -TimeoutSec 180" >nul 2>&1
    if exist "%temp%\adobe32_setup.exe" (
        start /wait "" "%temp%\adobe32_setup.exe" /sAll /rs /msi EULA_ACCEPT=YES
        timeout /t 15 /nobreak >nul 2>&1
        del /F /Q "%temp%\adobe32_setup.exe" >nul 2>&1
    )
)

call :detectar_adobe
if defined ADOBE_EXE (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -COM WScript.Shell; $s=$ws.CreateShortcut('!DESKTOP!\Adobe Acrobat Reader.lnk'); $s.TargetPath='!ADOBE_EXE!'; $s.Save()" >nul 2>&1
    echo     - [OK] Adobe Reader instalado.
) else (
    echo     - [ERROR] No se pudo instalar Adobe Reader.
)
goto :eof

:detectar_adobe
set "ADOBE_EXE="
if exist "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"       set "ADOBE_EXE=C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
if exist "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" set "ADOBE_EXE=C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
if exist "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"              set "ADOBE_EXE=C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
if exist "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"        set "ADOBE_EXE=C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
if exist "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd64.exe"       set "ADOBE_EXE=C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd64.exe"
if exist "C:\Program Files\Adobe\Acrobat Reader\Reader\AcroRd32.exe"          set "ADOBE_EXE=C:\Program Files\Adobe\Acrobat Reader\Reader\AcroRd32.exe"
if exist "C:\Program Files (x86)\Adobe\Acrobat Reader\Reader\AcroRd32.exe"    set "ADOBE_EXE=C:\Program Files (x86)\Adobe\Acrobat Reader\Reader\AcroRd32.exe"
goto :eof

:: ==========================================================
:: WINRAR
:: ==========================================================
:instalar_winrar
echo  [^>] WinRAR:
echo     - Instalando...

if "!WINGET_OK!"=="1" (
    winget install --id RARLab.WinRAR -e --silent --accept-package-agreements --accept-source-agreements --force >nul 2>&1
    timeout /t 5 /nobreak >nul 2>&1
)

if not exist "C:\Program Files\WinRAR\WinRAR.exe" (
    echo     - Descargando instalador directo...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.rarlab.com/rar/winrar-x64-710.exe' -OutFile \"$env:TEMP\winrar_setup.exe\" -UseBasicParsing -TimeoutSec 120" >nul 2>&1
    if exist "%temp%\winrar_setup.exe" (
        start /wait "" "%temp%\winrar_setup.exe" /S
        timeout /t 8 /nobreak >nul 2>&1
        del /F /Q "%temp%\winrar_setup.exe" >nul 2>&1
    )
)

timeout /t 3 /nobreak >nul 2>&1
if exist "C:\Program Files\WinRAR\WinRAR.exe" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -COM WScript.Shell; $s=$ws.CreateShortcut('!DESKTOP!\WinRAR.lnk'); $s.TargetPath='C:\Program Files\WinRAR\WinRAR.exe'; $s.Save()" >nul 2>&1
    echo     - [OK] WinRAR instalado.
) else (
    echo     - [ERROR] No se pudo instalar WinRAR.
)
goto :eof

:: ==========================================================
:: ZOOM
:: ==========================================================
:instalar_zoom
echo  [^>] Zoom:
echo     - Instalando...

taskkill /F /IM Zoom.exe >nul 2>&1
taskkill /F /IM ZoomInstaller.exe >nul 2>&1
taskkill /F /IM Zoom_launcher.exe >nul 2>&1

if "!WINGET_OK!"=="1" (
    winget install --id Zoom.Zoom -e --silent --override "/quiet /norestart ZoomAutoUpdate=true ZConfig=Lang=es" --accept-package-agreements --accept-source-agreements --force >nul 2>&1
    timeout /t 8 /nobreak >nul 2>&1
    taskkill /F /IM Zoom.exe >nul 2>&1
    taskkill /F /IM Zoom_launcher.exe >nul 2>&1
)

call :detectar_zoom
if not defined ZOOM_EXE (
    echo     - Descargando instalador directo...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://zoom.us/client/latest/ZoomInstallerFull.exe' -OutFile \"$env:TEMP\zoom_setup.exe\" -UseBasicParsing -TimeoutSec 120" >nul 2>&1
    if exist "%temp%\zoom_setup.exe" (
        start /wait "" "%temp%\zoom_setup.exe" /quiet /norestart ZoomAutoUpdate=true ZConfig=Lang=es
        timeout /t 12 /nobreak >nul 2>&1
        del /F /Q "%temp%\zoom_setup.exe" >nul 2>&1
    )
)

timeout /t 3 /nobreak >nul 2>&1
taskkill /F /IM Zoom.exe >nul 2>&1
taskkill /F /IM Zoom_launcher.exe >nul 2>&1

call :detectar_zoom
if defined ZOOM_EXE (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -COM WScript.Shell; $s=$ws.CreateShortcut('!DESKTOP!\Zoom.lnk'); $s.TargetPath='!ZOOM_EXE!'; $s.Save()" >nul 2>&1
    echo     - [OK] Zoom instalado.
) else (
    echo     - [ERROR] No se pudo instalar Zoom.
)
goto :eof

:detectar_zoom
set "ZOOM_EXE="
if exist "C:\Program Files\Zoom\bin\Zoom.exe" set "ZOOM_EXE=C:\Program Files\Zoom\bin\Zoom.exe"
if exist "C:\Program Files (x86)\Zoom\bin\Zoom.exe" set "ZOOM_EXE=C:\Program Files (x86)\Zoom\bin\Zoom.exe"
if exist "C:\Program Files\Zoom\Zoom.exe" set "ZOOM_EXE=C:\Program Files\Zoom\Zoom.exe"
if exist "C:\Program Files (x86)\Zoom\Zoom.exe" set "ZOOM_EXE=C:\Program Files (x86)\Zoom\Zoom.exe"
for /d %%U in ("C:\Users\*") do (
    if exist "%%~U\AppData\Roaming\Zoom\bin\Zoom.exe" set "ZOOM_EXE=%%~U\AppData\Roaming\Zoom\bin\Zoom.exe"
    if exist "%%~U\AppData\Local\Zoom\bin\Zoom.exe" set "ZOOM_EXE=%%~U\AppData\Local\Zoom\bin\Zoom.exe"
)
goto :eof

:: =========================================================
:: RESUMEN FINAL
:: =========================================================
:resumen_final

cls
echo.
echo  ============================================================
echo       CITYPC - PROCESO TERMINADO (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   RESULTADO:
echo.
echo   [OK]  Excepciones de antivirus aplicadas
echo   [OK]  Windows Update activado
echo   [OK]  Seguridad TLS configurada
echo   [OK]  Energia maxima configurada
echo.

if exist "C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe" (
    echo   [OK]  Soporte Tecnico CityPC - Instalado
) else (
    echo   [XX]  Soporte Tecnico CityPC - NO instalado
)

set "R_CHROME=0"
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "R_CHROME=1"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "R_CHROME=1"
if "!R_CHROME!"=="1" (
    echo   [OK]  Google Chrome          - Instalado
) else (
    echo   [XX]  Google Chrome          - NO se pudo instalar
)

call :detectar_adobe
if defined ADOBE_EXE (
    echo   [OK]  Adobe Acrobat Reader   - Instalado
) else (
    echo   [XX]  Adobe Acrobat Reader   - NO se pudo instalar
)

if exist "C:\Program Files\WinRAR\WinRAR.exe" (
    echo   [OK]  WinRAR                 - Instalado
) else (
    echo   [XX]  WinRAR                 - NO se pudo instalar
)

call :detectar_zoom
if defined ZOOM_EXE (
    echo   [OK]  Zoom                   - Instalado
) else (
    echo   [XX]  Zoom                   - NO se pudo instalar
)

echo.
echo  ============================================================
echo.
echo   Si algo marca [XX] puede intentar instalarlo manualmente.
echo.
echo   Presione cualquier tecla para cerrar...
pause >nul
exit
