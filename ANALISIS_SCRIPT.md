# Análisis técnico de `Instalador_CityPC.bat`

## Resumen ejecutivo

El script está diseñado como **instalador postformateo con auto-actualización** y flujo guiado para técnicos, pero también ejecuta acciones de alto impacto sobre seguridad del sistema (especialmente Microsoft Defender).

En términos funcionales:

- Valida ejecución como administrador.
- Se autoactualiza desde GitHub (`version.txt` + descarga del `.bat`).
- Aplica exclusiones y desactivaciones de Defender.
- Intenta activar Windows Update y TLS 1.2.
- Ajusta energía a alto rendimiento.
- Copia utilitarios desde USB (`Soporte...exe`, `Ninite.exe`).
- Muestra un resumen final y abre páginas/configuraciones.

## Hallazgos por área

### 1) Flujo general y robustez

**Positivo**
- El flujo está bien segmentado por etapas (`[1/7] ... [7/7]`) y con mensajes claros para usuario final.
- Tiene manejo de escenarios sin internet, sin bloquear toda la ejecución.

**Riesgos/limitaciones**
- Usa `goto` para ramificaciones múltiples; es válido en batch, pero complica mantenimiento cuando crezca.
- No hay bitácora persistente (archivo de log) para auditoría o soporte remoto.

### 2) Auto-update desde GitHub

**Positivo**
- Verifica conectividad y descarga `version.txt` antes de actualizar.
- Si falla descarga o parseo de versión, cae a modo seguro (`skip_update`).

**Riesgos críticos**
- No hay validación criptográfica del archivo descargado (firma/hash). Si el origen se compromete, el equipo ejecuta código arbitrario con privilegios administrativos.
- Se sobrescribe directamente el instalador en USB y luego se relanza elevado.

### 3) Seguridad / Defender

**Observación principal**
- El script no solo agrega exclusiones, también **desactiva protección en tiempo real y políticas de Defender**, incluyendo llaves de registro de política.

**Impacto**
- Deja el sistema en estado de exposición alta.
- Puede incumplir políticas corporativas, controles EDR y requisitos de cumplimiento.
- Varias rutas/procesos excluidos están relacionados con herramientas KMS, lo cual incrementa riesgo operativo y de reputación.

### 4) Windows Update y red

**Positivo**
- Intenta habilitar servicios relevantes (`wuauserv`, `bits`, `dosvc`) y dispara `usoclient`.
- Verifica conectividad con 8.8.8.8 y 1.1.1.1.

**Limitación**
- `usoclient` no siempre es confiable en todas las versiones/ediciones de Windows; puede no reflejar estado real de actualización.

### 5) TLS/certificados

**Positivo**
- Configura `SchUseStrongCrypto` y parámetros TLS 1.2 cliente.
- Intenta actualizar certificados raíz con `certutil -generateSSTFromWU`.

**Limitación**
- Sólo toca TLS 1.2 cliente; no contempla otros escenarios (p. ej. restricciones GPO o catálogos WSUS).

### 6) Instalación de software y soporte

**Positivo**
- Busca ejecutables en varias rutas de USB (resiliente a estructura variable).
- Crea acceso directo de soporte con icono cuando existe.

**Limitación**
- No valida integridad/firmas de ejecutables locales (`Ninite.exe`, `Soporte...exe`).
- El resumen final asume éxito de pasos previos aunque muchos comandos silencian error (`>nul 2>&1`).

## Riesgos prioritarios (Top 5)

1. **Desactivación extensa de Defender y notificaciones**.
2. **Auto-update sin verificación de firma/hash**.
3. **Ejecución privilegiada de binarios de USB sin validación**.
4. **Baja trazabilidad por ausencia de logging persistente**.
5. **Falsos positivos de éxito por ocultar errores sistemáticamente**.

## Recomendaciones concretas

1. **Restaurar seguridad al final**
   - Reactivar `RealtimeMonitoring`, `BehaviorMonitoring`, `IOAV`, `ScriptScanning`.
   - Evitar llaves de política permanentes de deshabilitación.

2. **Asegurar el auto-update**
   - Publicar hash SHA-256 firmado del `.bat` y verificar antes de reemplazar.
   - Considerar releases firmadas y validación de autor/origen.

3. **Validar binarios antes de ejecutar/copy**
   - `Get-AuthenticodeSignature` o hash esperado por versión.
   - Rechazar ejecución si firma no válida o hash no coincide.

4. **Agregar logging**
   - Registrar cada etapa en `C:\CityPC\logs\instalador.log` con timestamp.
   - Guardar códigos de salida de comandos críticos.

5. **Reducir superficie de exclusiones**
   - Mantener sólo exclusiones indispensables y temporales.
   - Documentar claramente motivo técnico por exclusión.

## Conclusión

El script es funcional como automatizador de preparación de equipos, pero su estado actual prioriza rapidez sobre seguridad. Para operación profesional y sostenible, conviene endurecer verificación de integridad, reducir desactivaciones de Defender y mejorar observabilidad (logs + validaciones).
