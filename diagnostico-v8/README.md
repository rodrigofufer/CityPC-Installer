# USB Diagnostico CityPC V8 unificado

Fuente del release pasivo `diagnostico-v8`. Se publica en GitHub para conservar y auditar el bundle, pero no activa ni instala V8 en las USB existentes. El indicador legacy de la raíz continúa en V7 hasta completar configuración privada, webhook v2 y validación Windows.

## Objetivo

JP y Urano ejecutan exactamente el mismo `Diagnostico_CityPC.bat` y el mismo
`usbdiag-wifi-readiness.ps1`. La unica diferencia permitida entre ambas USB es
`wifi.local.cmd`, que contiene exclusivamente `WIFI_SSID` y `WIFI_PASS`.

`usbdiag.shared.local.cmd` contiene el endpoint HTTPS y el token acotado. Debe
ser byte por byte identico en ambas USB y nunca se versiona.

## Archivos publicables

- `Diagnostico_CityPC.bat`: diagnostico comun y updater estricto.
- `usbdiag-wifi-readiness.ps1`: conexion y comprobacion acotada de red.
- `usbdiag-bundle-commit.ps1`: commit transaccional del bundle despues de que
  el BAT padre sale; hace backup, verifica, revierte completo y reinicia.
- `version_diagnostico.txt`: version comun, actualmente `8`.
- `update-manifest.json`: hashes SHA-256 del paquete comun.
- `update-channel.json.example`: contrato del puntero pequeno que se publica en
  `main`; el `update-channel.json` real apunta a un commit inmutable y al hash
  exacto de su manifiesto.
- ejemplos `.example`, documentacion y validadores.

## Archivos privados por USB

- `usbdiag.shared.local.cmd`: igual en JP y Urano.
- `wifi.local.cmd`: unico archivo que cambia por sucursal.

El BAT no contiene credenciales, tokens ni una URL productiva autenticada.

## Garantias V8

- una sola ejecucion por equipo mediante lock fail-closed;
- Wi-Fi no se declara listo hasta comprobar salida de `netsh`, SSID exacto,
  IPv4 valida, DNS y HTTPS;
- espera de asociacion acotada a 45 segundos y comandos `netsh` acotados;
- XML Wi-Fi temporal unico por ejecucion;
- perfil de conexion manual con nombre unico por ejecucion; nunca borra un
  perfil preexistente llamado como el SSID y retira solo su perfil temporal;
- DNS para GitHub y el endpoint, GET HTTPS del canal con contrato completo y
  handshake TLS autenticado contra el host del webhook sin enviar un POST de
  negocio;
- el updater descarga el canal con tres intentos y limite total de 35 segundos;
  luego descarga BAT, helper, committer y manifiesto desde un unico commit
  inmutable, con tres intentos por archivo, dos rutas GitHub ancladas al mismo
  commit y un limite total de 100 segundos;
- una version remota menor nunca hace downgrade; una version mayor actualiza y
  una version igual compara los tres hashes instalados para reparar drift,
  incluso un committer ausente o corrupto;
- el manifiesto debe cubrir exactamente los tres archivos comunes; se validan
  version interna, estructura, tamanos y hashes antes de mutar;
- el committer recibe el PID real del BAT y no copia nada hasta confirmar su
  salida; respalda los archivos existentes, registra tambien los ausentes,
  instala helper + committer + BAT, confirma los tres hashes y revierte el
  bundle completo si cualquier copia o comprobacion falla;
- la transaccion queda en la USB con fases `prepared`, `installing` e
  `installed-awaiting-confirmation`; tras un corte de energia la siguiente
  apertura la recupera o revierte antes de permitir el diagnostico;
- el lock guarda el PID de su propietario y solo elimina un lock viejo si ese
  proceso ya no existe; durante el handoff el committer toma la propiedad;
- la configuracion privada nunca entra al mapa actualizable;
- el reinicio `--updated` lleva token y version; no salta el updater hasta que
  el BAT nuevo vuelve a verificar estado, manifiesto y los tres hashes. La
  transaccion se conserva hasta esa confirmacion, evitando loops y falsos OK;
- contrato completo del pipeline antes de mostrar nube confirmada;
- ticket exactamente de cinco digitos.

## Limites deliberados

- El perfil XML necesita contener la clave en memoria y en un archivo temporal
  durante segundos; el helper lo elimina en `finally` y el BAT vuelve a limpiar.
- SHA-256 detecta corrupcion, pero no sustituye una firma criptografica si el
  repositorio fuera comprometido.
- Ningun updater puede prometer actualizar sin energia, internet, GitHub
  disponible y una USB escribible. En esos casos V8 usa reintentos acotados y
  continua con el bundle local ya confirmado; si encuentra una transaccion
  incompleta o un rollback no verificable, falla cerrado en vez de fingir una
  actualizacion.
- La conexion ocurre antes del fetch. Si falta el helper Wi-Fi o esta tan
  corrupto que no puede establecer y comprobar red, V8 falla cerrado y exige
  recopiado fisico; si el helper aun funciona pero su hash deriva, la
  comprobacion de version igual lo repara.
- El POST de negocio sigue siendo unico y fail-closed. No se reintenta de forma
  automatica hasta que el webhook v2 tenga idempotencia durable de punta a
  punta; ante una confirmacion incierta se conserva el reporte local y se
  indica no volver a ejecutar para evitar notas o mensajes duplicados.
- No se ejecutaron pruebas contra el webhook vivo durante esta publicación GitHub-only.
- La primera validacion de cada USB debe hacerse fisicamente con un endpoint de
  prueba no mutante antes de habilitar produccion.

## Validacion local

```bash
python3 validate_v8_static.py
python3 test_v8_bundle_updater.py
```

Para comparar dos USB montadas:

```bash
python3 validate_v8_static.py --pair /ruta/USB_URANO /ruta/USB_JP
```

La comparacion permite diferencias solamente en `wifi.local.cmd`; exige que el
BAT, helper, committer, manifiesto y configuracion compartida sean identicos.
