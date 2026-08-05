# Migracion V7 a V8 unificado

Este directorio conserva la fuente pasiva publicada de V8. No modifica el V7 activo, n8n, VPS ni las USB; la activación requiere los pasos de este documento.

## Cambio de arquitectura

V7 publica dos BAT con Wi-Fi y endpoint embebidos. V8 usa un solo BAT y separa
configuracion:

```text
Diagnostico_CityPC.bat              identico
usbdiag-wifi-readiness.ps1          identico
usbdiag-bundle-commit.ps1           identico
usbdiag.shared.local.cmd            identico y privado
wifi.local.cmd                      unico archivo distinto
```

## Orden seguro de corte

1. Crear y validar un webhook v2 autenticado sin retirar v1.
2. Provisionar fisicamente JP y Urano conforme a `INSTALACION.md`.
3. Pasar pruebas de fallo y exito en Windows real por sucursal.
4. Comparar ambas USB con `validate_v8_static.py --pair`.
5. Publicar primero un commit de release inmutable con BAT, helper, committer y
   manifiesto; calcular y conservar su SHA de commit y SHA-256 de manifiesto.
6. En un segundo commit, actualizar solamente `update-channel.json` en `main`
   para que apunte al commit de release y al hash exacto del manifiesto. Nunca
   construir version, manifiesto y archivos desde lecturas independientes de
   un `main` mutable.
7. Observar una primera ejecucion real por sucursal sin repetir webhooks.
8. Retirar v1 y rotar las credenciales Wi-Fi historicamente expuestas.

## Advertencia de compatibilidad V7

El updater V7 busca `Diagnostico_JP.bat` o `Diagnostico_Urano.bat`. Subir solo
`Diagnostico_CityPC.bat` y cambiar la version remota a 8 dejaria las copias V7
buscando sus nombres anteriores. Por ello la migracion inicial es fisica o
requiere un bootstrap transicional auditado; no se debe subir `version=8` antes
de resolver ese paso.

## Updater V8

Una vez instalado fisicamente, V8 obtiene un canal pequeno desde `main`. Ese
canal contiene version, commit inmutable de 40 caracteres y SHA-256 del
manifiesto. Todo el resto se descarga desde ese mismo commit: manifiesto,
`Diagnostico_CityPC.bat`, `usbdiag-wifi-readiness.ps1` y
`usbdiag-bundle-commit.ps1`. Antes de mutar exige manifiesto con exactamente
esos tres nombres, SHA-256, version interna, tamanos y estructura minima.

El canal se intenta tres veces durante hasta 35 segundos. Cada artefacto se
intenta tres veces, alternando dos rutas GitHub ancladas al mismo commit, con
un limite total de 100 segundos para el bundle. Una version menor se rechaza,
una mayor actualiza y una igual compara hashes para reparar un bundle derivado
o un committer ausente.

El committer descargado corre desde `%TEMP%`, toma propiedad del lock y espera
la salida del PID real del BAT padre. En la USB crea una transaccion durable,
respalda cada archivo existente y registra los ausentes, aplica helper +
committer + BAT, confirma sus hashes y, ante cualquier error, restaura y vuelve
a comprobar los tres. Un rollback no confirmado bloquea el diagnostico.
`usbdiag.shared.local.cmd` y `wifi.local.cmd` estan excluidos expresamente del
mapa actualizable.

El reinicio exitoso pasa `--updated <token> <version>`. El BAT nuevo no omite la
comprobacion hasta validar token, version, estado durable, manifiesto y los tres
hashes instalados. Solo entonces elimina la transaccion y evita una segunda
descarga en esa ejecucion. Una version igual solo reemplaza si detecta drift;
esto corrige el loop historico de comparar una version local fija contra otra
version remota que no correspondia al archivo descargado.

## Rollback

El updater conserva el rollback en `.usbdiag-update-transaction` hasta que el
BAT nuevo arranca y confirma. Si se corta la energia en `prepared`,
`installing`, `installed-awaiting-confirmation` o `confirmed`, la siguiente
apertura recupera o revierte antes de diagnosticar. Un staging sin `state.json`
se limpia porque el committer escribe estado antes de cualquier mutacion. El
lock con PID tambien se recupera automaticamente si su propietario ya no
existe.

Conservar ademas una imagen completa de cada USB antes del corte: esa imagen
cubre dano fisico, perdida total de la unidad o un fallo de escritura que haga
imposible restaurar. No reinyectar payloads ni repetir tickets durante rollback.
