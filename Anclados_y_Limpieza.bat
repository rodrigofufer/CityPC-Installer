@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
color 1F
title CITYPC - ANCLADOS Y LIMPIEZA
mode con: cols=100 lines=50

:: =========================================================
:: VERSION LOCAL
:: =========================================================
set "LOCAL_VER=47"
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
    echo       2. Busca el archivo "Anclados_y_Limpieza.bat"
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
echo        CITYPC - ANCLADOS Y LIMPIEZA (V%LOCAL_VER%)
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
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/Anclados_y_Limpieza.bat','%temp%\Anclados_y_Limpieza_update.bat') }catch{}" >nul 2>&1

if not exist "%temp%\Anclados_y_Limpieza_update.bat" (
    echo   [ERROR] No se pudo descargar. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Reemplazar archivo en la USB
copy /Y "%temp%\Anclados_y_Limpieza_update.bat" "%~dp0Anclados_y_Limpieza.bat" >nul 2>&1
del /F /Q "%temp%\Anclados_y_Limpieza_update.bat" >nul 2>&1

echo.
echo   [OK] Actualizado a V!REMOTE_VER!. Reiniciando...
echo.
timeout /t 2 /nobreak >nul 2>&1

:: Reiniciar CON permisos de administrador
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process cmd -ArgumentList '/c \"^"%~dp0Anclados_y_Limpieza.bat^\"' -Verb RunAs" >nul 2>&1
exit

:skip_update

title CITYPC - ANCLADOS Y LIMPIEZA (V%LOCAL_VER%)

cls
echo.
echo  ============================================================
echo        CITYPC - ANCLADOS Y LIMPIEZA (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   No cierre esta ventana. El proceso puede tardar unos segundos.
echo.
echo  ============================================================
echo.

:: =========================================================
:: 1. LIMPIAR MENU INICIO - RECOMENDADOS Y RECIENTES
:: =========================================================
echo  [1/3] Limpiando Menu Inicio (Recomendados y Recientes)...
echo.

:: Desactivar seguimiento de documentos recientes
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_TrackDocs /t REG_DWORD /d 0 /f >nul 2>&1
echo    [OK] Seguimiento de documentos recientes desactivado.

:: Desactivar aplicaciones recien agregadas
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_TrackProgs /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Explorer" /v HideRecentlyAddedApps /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Windows\Explorer" /v HideRecentlyAddedApps /t REG_DWORD /d 1 /f >nul 2>&1
echo    [OK] Aplicaciones recientes ocultadas.

:: Desactivar recomendaciones de Windows 11 en el menu inicio
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_IrisRecommendations /t REG_DWORD /d 0 /f >nul 2>&1
echo    [OK] Recomendaciones del menu inicio desactivadas.

:: Desactivar sugerencias y contenido promovido
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338388Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-353694Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-353696Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SoftLandingEnabled /t REG_DWORD /d 0 /f >nul 2>&1
echo    [OK] Sugerencias y contenido promovido desactivados.

:: Limpiar archivos recientes
del /F /Q "%AppData%\Microsoft\Windows\Recent\*" >nul 2>&1
del /F /Q "%AppData%\Microsoft\Windows\Recent\AutomaticDestinations\*" >nul 2>&1
del /F /Q "%AppData%\Microsoft\Windows\Recent\CustomDestinations\*" >nul 2>&1
echo    [OK] Historial de archivos recientes limpiado.

:: =========================================================
:: 2. LIMPIAR NOTIFICACIONES
:: =========================================================
echo.
echo  [2/3] Limpiando Notificaciones...
echo.

:: Crear script PowerShell para limpiar notificaciones
set "NOTIF_SCRIPT=%temp%\CityPC_LimpiarNotif.ps1"

> "%NOTIF_SCRIPT%" (
    echo $ErrorActionPreference = 'SilentlyContinue'
    echo.
    echo # Metodo 1: Limpiar via API de Windows Runtime
    echo try {
    echo     [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] ^| Out-Null
    echo     [Windows.UI.Notifications.ToastNotificationManager]::History.Clear^(''^)
    echo } catch {}
    echo.
    echo # Metodo 2: Limpiar base de datos de notificaciones
    echo try {
    echo     $nfPath = "$env:LOCALAPPDATA\Microsoft\Windows\Notifications"
    echo     Stop-Service -Name WpnUserService_* -Force -ErrorAction SilentlyContinue
    echo     Start-Sleep -Seconds 1
    echo     Get-ChildItem "$nfPath\wpndatabase*.db*" -ErrorAction SilentlyContinue ^| Remove-Item -Force -ErrorAction SilentlyContinue
    echo     Start-Service -Name WpnUserService_* -ErrorAction SilentlyContinue
    echo } catch {}
)

powershell -ExecutionPolicy Bypass -NoProfile -File "%NOTIF_SCRIPT%" >nul 2>&1
del /F /Q "%NOTIF_SCRIPT%" >nul 2>&1

echo    [OK] Notificaciones limpiadas.

:: =========================================================
:: 3. ANCLAR APLICACIONES A LA BARRA DE TAREAS
:: =========================================================
echo.
echo  [3/3] Anclando aplicaciones a la barra de tareas...
echo.

:: Detectar rutas de las aplicaciones
set "CHROME_EXE="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not defined CHROME_EXE if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not defined CHROME_EXE if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=%LocalAppData%\Google\Chrome\Application\chrome.exe"

set "WORD_EXE="
if exist "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" set "WORD_EXE=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if not defined WORD_EXE if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE" set "WORD_EXE=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
if not defined WORD_EXE if exist "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE" set "WORD_EXE=C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"
if not defined WORD_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE" set "WORD_EXE=C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
if not defined WORD_EXE if exist "C:\Program Files\Microsoft Office\Office15\WINWORD.EXE" set "WORD_EXE=C:\Program Files\Microsoft Office\Office15\WINWORD.EXE"
if not defined WORD_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE" set "WORD_EXE=C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE"

set "EXCEL_EXE="
if exist "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if not defined EXCEL_EXE if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
if not defined EXCEL_EXE if exist "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
if not defined EXCEL_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
if not defined EXCEL_EXE if exist "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files\Microsoft Office\Office15\EXCEL.EXE"
if not defined EXCEL_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE" set "EXCEL_EXE=C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE"

set "POWERPOINT_EXE="
if exist "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
if not defined POWERPOINT_EXE if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
if not defined POWERPOINT_EXE if exist "C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE"
if not defined POWERPOINT_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE"
if not defined POWERPOINT_EXE if exist "C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE"
if not defined POWERPOINT_EXE if exist "C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE" set "POWERPOINT_EXE=C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE"

:: Crear script PowerShell para anclar aplicaciones
set "PIN_SCRIPT=%temp%\CityPC_Anclar.ps1"
set "PIN_RESULT=%temp%\CityPC_PinResult.txt"

> "%PIN_SCRIPT%" (
    echo $ErrorActionPreference = 'SilentlyContinue'
    echo.
    echo function Pin-ToTaskbar {
    echo     param([string]$ExePath, [string]$AppName^)
    echo.
    echo     if (-not (Test-Path $ExePath^)^) {
    echo         return 'NOT_FOUND'
    echo     }
    echo.
    echo     # Intentar anclar usando el verbo del shell
    echo     try {
    echo         $shell = New-Object -ComObject Shell.Application
    echo         $dir = $shell.Namespace([System.IO.Path]::GetDirectoryName($ExePath^)^)
    echo         $item = $dir.ParseName([System.IO.Path]::GetFileName($ExePath^)^)
    echo         $verbs = $item.Verbs(^)
    echo         $pinVerb = $null
    echo         foreach ($v in $verbs^) {
    echo             if ($v.Name -match 'taskbar^|barra de tareas'^) {
    echo                 $pinVerb = $v
    echo                 break
    echo             }
    echo         }
    echo         if ($pinVerb^) {
    echo             $pinVerb.DoIt(^)
    echo             return 'PINNED'
    echo         }
    echo     } catch {}
    echo.
    echo     # Intentar anclar creando acceso directo en carpeta de anclados
    echo     try {
    echo         $pinnedFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    echo         if (Test-Path $pinnedFolder^) {
    echo             $lnkName = $AppName + '.lnk'
    echo             $lnkPath = Join-Path $pinnedFolder $lnkName
    echo             if (-not (Test-Path $lnkPath^)^) {
    echo                 $ws = New-Object -ComObject WScript.Shell
    echo                 $shortcut = $ws.CreateShortcut($lnkPath^)
    echo                 $shortcut.TargetPath = $ExePath
    echo                 $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($ExePath^)
    echo                 $shortcut.Save(^)
    echo                 return 'SHORTCUT'
    echo             } else {
    echo                 return 'ALREADY'
    echo             }
    echo         }
    echo     } catch {}
    echo.
    echo     return 'FAILED'
    echo }
    echo.
    echo # Anclar cada aplicacion
    echo $apps = @(
    echo     @{Path='%CHROME_EXE%'; Name='Google Chrome'},
    echo     @{Path='%WORD_EXE%'; Name='Microsoft Word'},
    echo     @{Path='%EXCEL_EXE%'; Name='Microsoft Excel'},
    echo     @{Path='%POWERPOINT_EXE%'; Name='Microsoft PowerPoint'}
    echo ^)
    echo.
    echo foreach ($app in $apps^) {
    echo     if ($app.Path -and $app.Path -ne ''^) {
    echo         $result = Pin-ToTaskbar -ExePath $app.Path -AppName $app.Name
    echo         "$($app.Name^)=$result" ^| Out-File -Append '%PIN_RESULT%' -Encoding UTF8
    echo     } else {
    echo         "$($app.Name^)=NOT_FOUND" ^| Out-File -Append '%PIN_RESULT%' -Encoding UTF8
    echo     }
    echo }
    echo.
    echo # Reiniciar explorer para aplicar cambios de anclado
    echo Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    echo Start-Sleep -Seconds 2
    echo if (-not (Get-Process explorer -ErrorAction SilentlyContinue^)^) {
    echo     Start-Process explorer.exe
    echo }
)

:: Limpiar resultado anterior
if exist "%PIN_RESULT%" del /F /Q "%PIN_RESULT%" >nul 2>&1

powershell -ExecutionPolicy Bypass -NoProfile -File "%PIN_SCRIPT%" >nul 2>&1
del /F /Q "%PIN_SCRIPT%" >nul 2>&1

:: Leer resultados y mostrar
set "CHROME_PINNED=0"
set "WORD_PINNED=0"
set "EXCEL_PINNED=0"
set "PPT_PINNED=0"

if exist "%PIN_RESULT%" (
    for /f "tokens=1,2 delims==" %%A in ('type "%PIN_RESULT%"') do (
        if "%%B"=="PINNED" (
            echo    [OK] %%A - Anclado a la barra de tareas.
            if "%%A"=="Google Chrome" set "CHROME_PINNED=1"
            if "%%A"=="Microsoft Word" set "WORD_PINNED=1"
            if "%%A"=="Microsoft Excel" set "EXCEL_PINNED=1"
            if "%%A"=="Microsoft PowerPoint" set "PPT_PINNED=1"
        ) else if "%%B"=="ALREADY" (
            echo    [OK] %%A - Ya estaba anclado.
            if "%%A"=="Google Chrome" set "CHROME_PINNED=1"
            if "%%A"=="Microsoft Word" set "WORD_PINNED=1"
            if "%%A"=="Microsoft Excel" set "EXCEL_PINNED=1"
            if "%%A"=="Microsoft PowerPoint" set "PPT_PINNED=1"
        ) else if "%%B"=="SHORTCUT" (
            echo    [OK] %%A - Acceso directo creado en barra de tareas.
            if "%%A"=="Google Chrome" set "CHROME_PINNED=1"
            if "%%A"=="Microsoft Word" set "WORD_PINNED=1"
            if "%%A"=="Microsoft Excel" set "EXCEL_PINNED=1"
            if "%%A"=="Microsoft PowerPoint" set "PPT_PINNED=1"
        ) else if "%%B"=="NOT_FOUND" (
            echo    [AVISO] %%A - No se encontro instalado.
        ) else if "%%B"=="FAILED" (
            echo    [AVISO] %%A - No se pudo anclar automaticamente.
        )
    )
    del /F /Q "%PIN_RESULT%" >nul 2>&1
) else (
    echo    [AVISO] No se pudo ejecutar el script de anclado.
)

:: Abrir aplicaciones que no se pudieron anclar para anclado manual
set "HAY_MANUALES=0"

if "!CHROME_PINNED!"=="0" if defined CHROME_EXE (
    set "HAY_MANUALES=1"
    start "" "!CHROME_EXE!"
    echo    [INFO] Chrome abierto para anclado manual.
)
if "!WORD_PINNED!"=="0" if defined WORD_EXE (
    set "HAY_MANUALES=1"
    start "" "!WORD_EXE!"
    echo    [INFO] Word abierto para anclado manual.
)
if "!EXCEL_PINNED!"=="0" if defined EXCEL_EXE (
    set "HAY_MANUALES=1"
    start "" "!EXCEL_EXE!"
    echo    [INFO] Excel abierto para anclado manual.
)
if "!PPT_PINNED!"=="0" if defined POWERPOINT_EXE (
    set "HAY_MANUALES=1"
    start "" "!POWERPOINT_EXE!"
    echo    [INFO] PowerPoint abierto para anclado manual.
)

:: =========================================================
:: RESUMEN FINAL
:: =========================================================
cls
echo.
echo  ============================================================
echo       CITYPC - ANCLADOS Y LIMPIEZA TERMINADO (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   RESULTADO:
echo.
echo   [OK]  Menu Inicio - Recomendados y recientes limpiados
echo   [OK]  Menu Inicio - Sugerencias desactivadas
echo   [OK]  Notificaciones limpiadas
echo.

if "!CHROME_PINNED!"=="1" (
    echo   [OK]  Google Chrome          - Anclado a barra de tareas
) else if defined CHROME_EXE (
    echo   [!!]  Google Chrome          - Abierto para anclar manualmente
) else (
    echo   [XX]  Google Chrome          - No encontrado
)

if "!WORD_PINNED!"=="1" (
    echo   [OK]  Microsoft Word         - Anclado a barra de tareas
) else if defined WORD_EXE (
    echo   [!!]  Microsoft Word         - Abierto para anclar manualmente
) else (
    echo   [XX]  Microsoft Word         - No encontrado
)

if "!EXCEL_PINNED!"=="1" (
    echo   [OK]  Microsoft Excel        - Anclado a barra de tareas
) else if defined EXCEL_EXE (
    echo   [!!]  Microsoft Excel        - Abierto para anclar manualmente
) else (
    echo   [XX]  Microsoft Excel        - No encontrado
)

if "!PPT_PINNED!"=="1" (
    echo   [OK]  Microsoft PowerPoint   - Anclado a barra de tareas
) else if defined POWERPOINT_EXE (
    echo   [!!]  Microsoft PowerPoint   - Abierto para anclar manualmente
) else (
    echo   [XX]  Microsoft PowerPoint   - No encontrado
)

echo.

if "!HAY_MANUALES!"=="1" (
    echo  ============================================================
    echo.
    echo   ANCLADO MANUAL: Algunas aplicaciones se abrieron porque
    echo   no se pudieron anclar automaticamente.
    echo.
    echo   Para anclarlas:
    echo     1. Busque el icono de la aplicacion en la barra de tareas
    echo     2. Dele clic DERECHO sobre el icono
    echo     3. Seleccione "Anclar a la barra de tareas"
    echo     4. Cierre la aplicacion cuando termine
    echo.
)

echo  ============================================================
echo.
echo   Presione cualquier tecla para cerrar...
pause >nul
exit
