# Instalacion fisica controlada

No copiar este candidato a una USB operativa hasta tener preparado el webhook
v2 autenticado y un endpoint de prueba que no cree tickets, notas ni mensajes.

## Preparar el paquete comun

1. Copiar a ambas USB todos los archivos del paquete V8.
2. Confirmar que no existan `Diagnostico_JP.bat` ni
   `Diagnostico_Urano.bat`; el unico lanzador es
   `Diagnostico_CityPC.bat`.
3. Confirmar que junto al BAT existan `usbdiag-wifi-readiness.ps1` y
   `usbdiag-bundle-commit.ps1`; ambos forman parte obligatoria del bundle.
4. Copiar `usbdiag.shared.local.cmd.example` como
   `usbdiag.shared.local.cmd` y completar los dos valores privados.
5. Copiar exactamente ese mismo archivo compartido a la otra USB.
6. Copiar `wifi.local.cmd.example` como `wifi.local.cmd` y completar solamente
   el SSID y password correspondientes a cada sucursal.
7. Si un valor CMD contiene `%`, escribirlo como `%%`. No usar saltos de linea
   ni comillas dobles dentro de los valores.

## Validacion antes de abrir el BAT

1. Ejecutar `python3 validate_v8_static.py` en la fuente del paquete.
2. Montar ambas USB y ejecutar la comparacion `--pair` indicada en README.
3. Confirmar que el unico archivo distinto sea `wifi.local.cmd`.
4. Confirmar que `usbdiag.shared.local.cmd` no sea un ejemplo ni contenga
   placeholders.
5. Guardar los hashes del BAT y helper en el acta de aprovisionamiento sin
   copiar credenciales; guardar tambien el hash del committer.

## Prueba fisica por sucursal

1. Usar una computadora de laboratorio, no un ticket real improvisado.
2. Ejecutar una sola instancia del BAT.
3. Confirmar que no aparece `[OK]` antes de SSID exacto, IPv4, DNS y HTTPS.
4. Probar credencial incorrecta, sin cobertura, sin DHCP y DNS bloqueado; cada
   caso debe fallar con un mensaje distinto y sin iniciar diagnostico.
5. Probar conexion correcta contra el endpoint v2 de prueba no mutante.
6. Al cerrar, comprobar que el perfil Wi-Fi temporal ya no existe.
   El perfil se llama `CityPC-Diagnostico-<nonce>`; un perfil previo llamado
   como el SSID no debe borrarse ni modificarse.
7. Repetir por separado en JP y Urano.
8. En una copia de laboratorio con version menor, probar una actualizacion
   completa. Confirmar una sola reapertura con `--updated <token> <version>`,
   tres hashes finales iguales al manifiesto, transaccion eliminada tras la
   confirmacion y cero segunda descarga.
9. Forzar en laboratorio un fallo de copia y confirmar que los tres archivos
   regresan a sus hashes anteriores; las configuraciones privadas deben quedar
   byte por byte iguales.
10. Borrar solamente el committer en la copia de laboratorio, mantener version
    igual a la remota y confirmar que la comprobacion de hashes lo repone sin
    exigir que el committer instalado exista antes del fetch.
11. Simular 2 fallos de descarga y confirmar exito al tercer intento; simular
    3 fallos y confirmar salida acotada, sin loop ni archivos parcialmente
    instalados.
12. Mantener vivo el BAT padre mas de 60 segundos en el punto de handoff y
    confirmar que el committer no libera el lock ni abre otra copia.
13. Cortar energia en laboratorio, por separado, despues de `prepared`, durante
    `installing` y antes de la confirmacion. En cada caso, abrir una sola vez y
    confirmar recuperacion o rollback completo antes de pedir ticket.

## Publicacion del canal de actualizacion

1. No cambiar `update-channel.json` todavia. Crear un commit de release con el
   BAT, helper, committer y `update-manifest.json` ya validados.
2. Registrar el SHA de commit de 40 caracteres y calcular el SHA-256 del
   manifiesto exactamente como quedo en ese commit.
3. Crear `update-channel.json` desde el ejemplo con schema, version numerica,
   ese commit y ese hash. Validar JSON y publicar el puntero en un segundo
   commit de `main`.
4. Descargar de forma independiente el canal; luego descargar los cuatro
   artefactos desde el commit indicado y volver a comprobar el hash del
   manifiesto y los tres hashes de archivos.
5. Probar primero con una USB de laboratorio y despues con una USB por sucursal.
   No mover el canal a otra release hasta terminar esa observacion.

No publicar un canal que apunte al propio commit del canal: el release debe ser
un commit inmutable anterior y el puntero se actualiza despues. Tampoco usar
`version_diagnostico.txt` y archivos de `main` como lecturas independientes.

## Lock de ejecucion

El lock vive en `%TEMP%\citypc_usbdiag_v8.lock` y contiene `owner.pid`. El BAT y
el committer transfieren esa propiedad. Si Windows o el proceso se cierra de
forma abrupta, la siguiente apertura solo retira el lock cuando ese PID ya no
existe. Si el PID sigue vivo, falla cerrado para no duplicar procesos; antes de
cualquier retiro manual, confirmar en Administrador de tareas que no exista
otra ejecucion del diagnostico o PowerShell asociado.

## No entregar un equipo cuando

- el BAT reporta perfil Wi-Fi pendiente;
- la nube no confirma el contrato completo;
- hay dos instancias abiertas;
- falta cualquier archivo comun;
- los hashes no coinciden;
- el ticket esta cerrado, resuelto o cancelado en el sistema vivo.
