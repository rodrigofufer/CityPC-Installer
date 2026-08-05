# CityPC Preparación V49 — actualización robusta

Estado del paquete: release inmutable `preparacion-v49`. La activación de `main` se hace únicamente mediante las fases A/B/C documentadas aquí. No implica que una USB física ya haya sido probada.

## Problema confirmado

- `Instalador_CityPC.bat` estaba en V48.
- `Anclados_y_Limpieza.bat` estaba en V47.
- Ambos consultaban `version.txt=48` y comparaban sólo por desigualdad.
- El instalador V48 no buscaba actualización; Anclados V47 descargaba un archivo y podía anunciar la versión remota sin comprobar que el BAT descargado realmente la declarara.
- Ambos reemplazaban el BAT en ejecución, sin SHA-256, rollback, lock ni sentinel de confirmación.
- `Reactivar_OneDrive.bat` no tenía versión ni autoactualización.

## Arquitectura propuesta

`preparacion-channel-v2.json` en `main` sólo descubre la versión. Apunta al tag inmutable `preparacion-v49` y fija el SHA-256 de `preparacion-bundle-v2.json`. Todos los artefactos ejecutables se descargan desde ese tag; nunca desde `main` mutable.

El bundle contiene y versiona por separado:

- `Instalador_CityPC.bat` V49;
- `Anclados_y_Limpieza.bat` V49;
- `Reactivar_OneDrive.bat` V2;
- `CityPC_Updater.ps1` V1.

Ejecutar cualquiera de los tres BAT repara/actualiza los cuatro archivos como una sola transacción. Se descargan todos a temporales únicos, se valida manifiesto, nombre, tamaño, `LOCAL_VER` y SHA-256, y sólo después se cierra el caller para reemplazar.

## Guardas

- Sólo actualiza hacia versiones mayores; una versión local legible mayor bloquea downgrade.
- Un BAT truncado o sin `LOCAL_VER` se repara desde el bundle; el OneDrive legado se reconoce sólo por su SHA-256 exacto.
- Lock global acotado a 5 segundos.
- Ventana total de red de 90 segundos; cada descarga tiene dos intentos de 12 segundos. Sin red o GitHub disponible, continúa la versión actual sin reemplazar nada.
- Backups durables y de transacción se verifican por SHA-256. Si la copia `previous` existe pero está corrupta, el rollback usa el backup durable sólo después de verificar su hash exacto.
- Commit de los cuatro archivos con rollback total verificado ante cualquier fallo.
- El BAT nuevo reinicia una sola vez con sentinel de 32 caracteres. Debe confirmar los hashes completos dentro de 60 segundos.
- Un resume inválido sale con código 30 y no ejecuta la preparación. Un estado ambiguo sale con 31.
- Si falta confirmación, el committer revierte, bloquea exactamente ese bundle 15 minutos y relanza la última copia conocida. Un bundle más nuevo no queda bloqueado.
- Tras corte eléctrico: `staged` se reanuda; `committing` se revierte; `committed` íntegro se confirma. Dos transacciones activas se ponen en cuarentena y no se elige una al azar.
- Si el BAT original no termina en 45 segundos, el committer marca el intento fallido y no relanza ninguna copia mientras el proceso original siga vivo; esto evita ejecuciones duplicadas.

SHA-256 valida integridad, no reemplaza una firma criptográfica. El canal y el repositorio GitHub siguen siendo la raíz de confianza.

## Transición sin loop

No publicar todo de una vez. La migración segura requiere tres fases y observación entre ellas.

1. Crear primero el tag inmutable `preparacion-v49` con los cuatro artefactos y el manifiesto exactos.
2. Phase A en `main`: conservar `version.txt=48`, publicar el channel y sólo el puente `transition/phase-a-main/Anclados_y_Limpieza.bat` V48. El Anclados V47 legado lo descarga una vez; el puente obtiene el helper desde el tag por hash y migra el bundle completo.
3. Validar en Windows real y en una USB de cada sucursal que Anclados V48 pasa a bundle V49, reinicia una vez y la segunda ejecución no actualiza.
4. Phase B: publicar los V49 y el channel de `transition/phase-b-preposition-main/`, pero conservar `version.txt=48`. Esperar propagación y verificar por raw, varias veces, que los tres BAT V49 y el helper/manifiesto del tag devuelven los hashes esperados. Así se evita que el cliente legado vea primero el número 49 y todavía descargue un BAT V48 desde CDN.
5. Phase C: cambiar únicamente `version.txt` a 49 usando `transition/phase-c-activate-main/version.txt`. Esto alcanza los USB donde sólo se ejecutaba Instalador V48 y los deja en el mismo bundle.
6. Mantener rollback del tag/canal y los backups físicos. No borrar `version.txt` hasta que no queden clientes del actualizador legado.

## Validación local

```bash
python3 validate_update_v2.py
python3 -m unittest -v tests/test_updater_protocol.py
git diff --check
```

La publicación GitHub-only fue autorizada con validación estática y simulada. Sigue pendiente validar en Windows real: PowerShell 5.1, NTFS y exFAT/FAT32, rutas con espacios, USB sin internet, corte durante cada movimiento, UAC, sentinel, rollback y segunda ejecución sin repetición. No declarar esas pruebas como realizadas hasta ejecutarlas físicamente.
