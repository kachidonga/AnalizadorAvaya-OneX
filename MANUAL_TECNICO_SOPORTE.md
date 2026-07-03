# Manual Técnico / Soporte — Analizador Avaya OneX Agent (V22)

> Documento de referencia interna para el personal de Soporte. Explica **qué hace
> la herramienta por dentro y cómo lo hace**: arquitectura, el pipeline de
> análisis, las fuentes de datos, la lógica de detección de cada interpretación y
> los puntos delicados a tener en cuenta al diagnosticar o modificar el script.
>
> Archivo activo: **`AnalizadorOneXV22.ps1`** (~4,300 líneas, PowerShell 5.1 + WinForms).
> Para el uso operativo de la GUI, ver [`MANUAL_USUARIO.md`](MANUAL_USUARIO.md).

---

## 1. Arquitectura general

Es un **único script PowerShell** que levanta una GUI WinForms y procesa logs en
memoria. No hay base de datos ni dependencias externas: todo el estado vive en
*hashtables* durante la corrida de **ANALIZAR**.

```
┌─────────────────────────── GUI (WinForms) ───────────────────────────┐
│  Fila de botones: IP · 1.BUSCAR · 2.ANALIZAR · 3.RUTA MANUAL · ...     │
│  Grid principal (DataGridView, 10 columnas)                           │
│  Menú contextual (diagnósticos) · Clic en celda (evidencia/aislar)    │
└───────────────────────────────────────────────────────────────────────┘
                              │
                              ▼  (botón ANALIZAR)
        ┌──────────────── MOTOR DE ANÁLISIS ────────────────┐
        │  PASO 1  ContactLog.xml   (fuente primaria)        │
        │  PASO 2  AudioLog.txt                              │
        │  PASO 3  IspeacLog.txt    (calidad)                │
        │  Pre-pasada MidCall  (teléfonos mid-call)          │
        │  PASO 4  OneXAgent.log    (conexiones, holds)      │
        │  PASO 5  EndpointLog.txt  (sesiones, fin, firma)   │
        │  Limpieza post-PASO5  (consultas, huérfanas)       │
        │  Barridos  (FALLBACK FIN · HOLD implícito · Zombie)│
        │  PASO 6  Render "Cerebro Forense" → Grid           │
        └────────────────────────────────────────────────────┘
```

### Estructura de datos central
Todo se acumula en `$EventosTiempo` — una hashtable cuya **clave es la hora**
(slot) y cuyo valor es un objeto con la interpretación y la evidencia de cada
fuente:

```
$EventosTiempo[$slot] = @{
    Interpretacion / Agente / Audio / Aux / Ispeac / SysLog / AppLog   # texto visible
    RawInterpretacion / RawAgente / ...                                 # log crudo (Tag)
    Sesion="-"; Tel="-"; ViId=""; Topic=""
    ColorInterpretacion / ColorAgente / ...                             # color por celda
}
```

**Dos tipos de slot (clave):**
- **Slot base** = `HH:mm:ss` (solo segundos). Lo que aporta el XML.
- **Slot ms** = `HH:mm:ss,fff` (con milisegundos). Lo crean PASO 4/5 para ordenar
  con precisión eventos que caen en el mismo segundo (p. ej. señal + INICIO).

El render del PASO 6 ordena las claves y vuelca cada slot no vacío a una fila.

---

## 2. Modos de carga y conexión

| Modo | Mecánica |
|------|----------|
| **Remoto (IP)** | `Test-Connection` (ping) → `New-PSDrive` montando `\\IP\c$` con `Get-Credential` → enumera `C:\Users` excluyendo perfiles del sistema. Loop que reintenta credenciales si fallan. La unidad se monta/desmonta puntualmente en cada operación. |
| **Ruta Manual** | `FolderBrowserDialog`; auto-detecta el rango de fechas leyendo las primeras 50 líneas de cada `EndpointLog.txt*` y posiciona el `DateTimePicker` en la fecha más reciente. |

`$Script:DirFinalGlobal` guarda la carpeta de logs resuelta; `$Script:RutaManual`
distingue el modo (vacío = remoto). Varios módulos (Extraer/Búsqueda) **reconectan**
la unidad si están en modo remoto.

### Filtro de fecha (crítico)
Los logs de Avaya usan formato **MM/dd/yyyy (gringo)**. El filtro `$FechaOmni`
acepta solo:
- `MM/dd/yyyy` (con ceros), `M/d/yyyy` (sin ceros) e ISO `yyyy-MM-dd`.

Se **excluyen** `dd/MM` y `d/M` a propósito, para evitar colisiones (ej.
`05/06/2026` interpretado como 5-jun vs 6-may). **Mantener este criterio** en
cualquier módulo que filtre por fecha — la búsqueda rápida y la extracción RAW lo
replican.

---

## 3. El pipeline de ANALIZAR, paso a paso

### PASO 1 — ContactLog.xml (fuente primaria)
Es la **fuente de verdad** de llamadas y números. Se parsea como `[xml]` y de cada
`ContactLogItem` tipo `Voice` se extraen: hora (`CreateTime` → hora local),
número (`Uri`), ID interno (`Id`), y dirección.

- Llamada **saliente** → se siembra en el slot base como INICIO/LÍNEA ABIERTA
  provisional; PASO 4/5 lo refinan.
- Llamada **entrante** → siembra el momento de *alerting*.
- Pobla `$Script:ModoContestacion` (VI_ID → modo) para la validación Auto-Answer.

> Si el XML está cargado, manda; los handlers de *alerting* del PASO 4 actúan solo
> como *fallback* (estilo V21) cuando el XML no está.

### PASO 2 — AudioLog.txt
Eventos de dispositivos/mezclador de audio (`AudioMixer`, etc.) → columna **Audio.log**.

### PASO 3 — IspeacLog.txt
Métricas de **calidad de red/voz** → columna **Calidad Red/Voz** (oculta por
defecto; la casilla *Ver Calidad de Red* la muestra junto con filas "Llamada en
curso (Analizando Calidad)").

### Pre-pasada MidCall
Antes del PASO 4 detecta **teléfonos actualizados a mitad de llamada** en
EndpointLog y los guarda en `$Script:MidCallPhones`, para que el PASO 4 **suprima**
falsos INICIOs cuando solo cambió el número de la misma llamada (transferencias
inter-agente).

### PASO 4 — OneXAgent.log (`one-X*.log`)
El más pesado. Lee todos los segmentos rotados (`.log`, `.log.1` … `.log.12`).
Aquí vive la mayor parte de la lógica de detección:

| Evento en el log | Qué produce |
|------------------|-------------|
| `OldState=New,NewState=Alerting` (entrante) | Da `cxtUUID` + `ConnectionId`; llena `$Script:CxtToConnId`; crea/ordena el slot de señal |
| `OldState=Ringing,NewState=Active` (saliente) | Da `cxtUUID`+`ConnectionId` para mostrar Sesión en INICIO Saliente |
| `PhoneService_CallStateChanged State=Alerting,Outgoing=False` | Puente **CM Auto-Answer**: marca señal entrante / reetiqueta saliente→entrante |
| `Setting Vi.InboundAcd=true` | **Veredicto real de ACD**: reescribe el tipo de la señal a ACD (no se confía solo en el número) |
| `VoiceInteractionImpl type=Active` (**PRIMARY CONNECTED**) | INICIO DE LLAMADA real (entrante/saliente) en slot ms, con *Ring Time* exacto vía UUID |
| `InnerState=CONNECTED` (**SECONDARY**) | Conexión secundaria; usa ID numérico, con *guards* anti-duplicado |
| `InnerState=BRIDGED_CONNECTED` + `RemoteParty=[,]` | `[!] Sesión bridge del sistema` (auto-transferencia, no acción del asesor) |
| `EnterReadyHandler` / Auto-In | "Asesor se cambia a Disponible (Confirmado por clic / Auto-In)" |
| `PressLineAppearance` | Apertura de línea (foquito) |
| `OldState=Active,NewState=Inactive` **sin** `OnRequestHoldSession` | Hold implícito → `$HoldImplicito` (lo materializa un barrido) |
| `WI.ADD` (`VoiceInteractionListImpl.Add`) | Tercera fuente de detección de señal (INTERNA/ACD/EXTERNA) |
| `Call ended` (del día filtrado) | `$MapeoTelP4` (sesId→tel), *fallback* de teléfono |
| `MakeCall` | Respaldo de dirección saliente |

#### Validación de dirección: el evento `Call StateChanged` (el "semáforo" de la llamada)

> **Este es el mecanismo central de validación de dirección del analizador.** El
> par de eventos `Call StateChanged` del CM (Communication Manager) es lo que nos
> permite afirmar con certeza si una llamada fue **entrante** o **saliente**, y no
> depender de heurísticas frágiles.

El conmutador emite una transición de estado por cada llamada. Nos interesan dos:

| Línea en `OneXAgent.log` | Dirección | Para qué la usamos |
|--------------------------|-----------|--------------------|
| `Call StateChanged Call[Id=<UUID>,ConnectionId=<N>,OldState=New,NewState=Alerting]` | **Entrante** | El teléfono está timbrando una llamada que **entra** |
| `Call StateChanged Call[Id=<UUID>,ConnectionId=<N>,OldState=Ringing,NewState=Active]` | **Saliente** | Una llamada que el agente **marcó** acaba de conectar |

**Por qué `OldState=New,NewState=Alerting` es la pieza clave (caso CM Auto-Answer):**

1. Con **"Se requiere uso de respuesta automática CM"** activo, el CM **contesta la
   llamada a nivel de conmutador** antes de que el agente toque nada.
2. En ese escenario, el `WI.ADD` (`VoiceInteractionListImpl.Add`) **nunca pasa por
   `state=Alerting`**: salta directo a `type=Active`. Por eso `$Script:AlertingHoras`
   **no se puebla**, y si nos guiáramos solo por eso, el `PRIMARY_CONNECTED`
   **clasificaría la llamada como Saliente por error**.
3. En cambio, `OldState=New,NewState=Alerting` **sí aparece siempre que el teléfono
   timbra**, para **todas** las llamadas ACD. Es la señal confiable que sobrevive al
   Auto-Answer.

**Qué hace el handler (líneas ~971–1004) cuando detecta `New→Alerting`:**

- Captura el **callUUID** (`Id=`) que más tarde reaparece como `cxt=` en el
  `PRIMARY_CONNECTED` → así se **reconstruye el puente** entre el timbrado y la
  conexión real.
- Crea la **"Agente con señal de llamada (Entrante - ACD)"** en un **slot ms**
  (en dorado), pero **solo si `WI.ADD` no la creó ya** en ese segundo (chequea slot
  base + ms-slots para no duplicar la señal del XML).
- Registra `$ConnIdToSeñalSlot[sesId] = slot` y marca `$DirLlamada[sesId] = "ENTRANTE"`.
- Pobla `$Script:AlertingPorCxtUuid[UUID]` (puente CM Auto-Answer hacia el INICIO).
- Mapea **`cxtUUID → ConnectionId` numérico** en `$Script:CxtToConnId`, que sirve para:
  - **bloquear SECONDARY tardíos** del mismo CallID (p. ej. un `Active→Active`
    inter-agente que llega 20–30 s después) y evitar un INICIO duplicado, y
  - **rellenar el campo `Sesión`** del INICIO DE LLAMADA (el `PRIMARY_CONNECTED`
    recupera el ConnectionId desde aquí).

> En resumen: `OldState=Ringing,NewState=Active` confirma **salientes**, y
> `OldState=New,NewState=Alerting` confirma **entrantes** (y es el único testigo
> fiable cuando el CM responde en automático). El diagnóstico de menú
> **"¿Por qué es Entrante o Saliente?"** se apoya directamente en esta evidencia.

### PASO 5 — EndpointLog.txt
Corre **después** del PASO 4 (orden importante). Aporta:
- `UpdateHistoryRecord` → `$MapeoTel` (sesId→tel, **autoritativo**).
- `ProcessSessionEndedEvent` → FIN de llamada, con clasificación según
  `OnRequestEndSession` (manual vs normal), `CUELGUE MANUAL`, `¡EVASIÓN!`.
- HOLD/UNHOLD manual (`OnRequestHoldSession`), Mute on/off.
- Intentos de **firma** ("Usuario intentando firmarse en la Ext. N" /
  "Recuperado por PBX").
- Consultas previas a conferencia/transferencia (`ConsultaConf` / `ConsultaTransf`).

> **Por qué el orden PASO 4 → PASO 5:** el FIN definitivo (manual vs normal) depende
> de `OnRequestEndSession`, que vive en EndpointLog. El PASO 4 detecta el cierre
> técnico pero no escribe el FIN; lo deja para que el PASO 5 lo clasifique.

### Limpieza post-PASO5
- ① Reetiqueta sesiones marcadas como **consulta** (CONSULTA DE TRANSFERENCIA/CONFERENCIA).
- ② Elimina filas **LÍNEA ABIERTA huérfanas** del PASO 4 cuya misma sesión ya tiene
  un INICIO dentro de los ~10 s siguientes (upgrade tardío del PASO 5).
- Limitación conocida: si el PASO 4 creó la fila con `Sesion="-"` (distinto
  segundo), no se puede identificar por sesión y queda fuera del barrido.

### Barridos finales
- **FALLBACK FIN** — aplica FIN a sesiones que `ProcessSessionEndedEvent` no cubrió,
  usando `$CallStateDisconnected` (`OldState=Active,NewState=Disconnected`).
- **HOLD implícito** — materializa "Hold automático del sistema (Sesión N)" desde
  `$HoldImplicito`.
- **Zombie detector** — limpia slots/sesiones inconsistentes.

### PASO 6 — Cerebro Forense (render)
Ordena las claves de `$EventosTiempo`, descarta slots vacíos o absorbidos
(`Interpretacion=""` + marcador `[ABSORBIDO:...]`), aplica colores por celda y
vuelca cada slot a una fila del grid. Guarda el log crudo en el `.Tag` de cada
celda para la evidencia al hacer clic.

---

## 4. Catálogo de interpretaciones y su disparador

| Interpretación (columna Actividad) | Fuente / disparador |
|------------------------------------|---------------------|
| Agente con señal de llamada (Entrante - INTERNA/ACD/EXTERNA - tel) | WI.ADD / Alerting; tipo ACD confirmado por `Vi.InboundAcd=true` |
| INICIO DE LLAMADA (Entrante) + Ring Time | PRIMARY CONNECTED (`type=Active`), UUID→tiempo de timbrado |
| INICIO DE LLAMADA (Saliente) | PRIMARY CONNECTED saliente / XML / `OldState=Ringing,NewState=Active` |
| LÍNEA ABIERTA SIN MARCAR | Línea levantada sin número |
| INICIO DE SESIÓN (Llamada Interna — Extensión N) | EndpointLog, sesión interna |
| Asesor se cambia a Disponible (Auto-In / Confirmado por clic) | `EnterReadyHandler` / `Invoke(fnu=auto-in)` |
| Extensión en línea y conectada a dispositivos de audio | Arranque del cliente con audio OK |
| Usuario intentando firmarse en la Ext. N / Recuperado por PBX | EndpointLog (firma) |
| HOLD MANUAL (Clic del Agente) | `OnRequestHoldSession` en EndpointLog |
| Hold automático del sistema (Sesión N) | `OldState=Active,NewState=Inactive` sin OnRequestHold (barrido) |
| Mute activado / desactivado | EndpointLog |
| CONSULTA DE CONFERENCIA / DE TRANSFERENCIA → número | Consulta previa detectada en PASO 5 |
| CONFERENCIA ESTABLECIDA / TRANSFERENCIA COMPLETADA | EndpointLog (habilita el menú *Ver detalle*) |
| FIN DE LLAMADA NORMAL | `ProcessSessionEndedEvent` sin marca manual |
| FIN DE LLAMADA MANUAL (Colgada por el Asesor) | `OnRequestEndSession` (manual) |
| CUELGUE MANUAL (Línea abierta sin marcar) | Cierre de línea abierta sin número |
| ¡EVASIÓN! Línea abandonada sin marcar (Timeout/Tapón) | Línea abierta abandonada |
| Número de cliente actualizado (Transferencia inter-agente) | MidCall: cambió el tel de la misma sesión |
| [!] Sesión bridge del sistema (auto-transferencia) | `BRIDGED_CONNECTED` + `RemoteParty=[,]` |
| [CM Auto-Answer] (sufijo) | Puente PhoneService Auto-Answer del Communication Manager |

---

## 5. Funciones de diagnóstico (menú contextual)

Se habilitan según el contenido de la fila (ver lógica en `Add_MouseDown`). Cada
una abre un reporte **paso a paso** que reconstruye el razonamiento sobre un
**snapshot** de los datos (`$Script:SnapEventosTiempo`):

| Opción | Sobre qué fila | Qué reconstruye |
|--------|----------------|-----------------|
| ¿Por qué no hay INICIO DE LLAMADA? | "señal de llamada" | Sigue el UUID: slot → `AlertingHoras` → INICIO; muestra dónde se rompió el vínculo |
| ¿Por qué no hay FIN DE LLAMADA? | INICIO / LÍNEA ABIERTA | Busca el cierre de sesión esperado |
| ¿Por qué finalizó esta llamada? | FIN / CUELGUE | Explica la clasificación (manual/normal/evasión) |
| ¿Por qué es Entrante o Saliente? | INICIO DE LLAMADA | Muestra la evidencia de dirección |
| Ver detalle: ¿Cómo se estableció la Conferencia? | CONFERENCIA ESTABLECIDA | DragDrop, consultas y participantes |
| Ver detalle: ¿Cómo se completó la Transferencia? | TRANSFERENCIA COMPLETADA | Líneas de log y duración del proceso |

> Estos reportes son la mejor herramienta cuando un caso "se ve raro": muestran
> exactamente qué evidencia encontró (o no) el motor y en qué paso.

---

## 6. Acciones de clic en celda (`Add_CellClick`)

| Clic en… | Resultado |
|----------|-----------|
| Cualquier celda regular (Actividad/Endpoint/Audio/Aux/Sys/App) | MessageBox con el **log crudo** guardado en `.Tag` (evidencia) |
| Columna **Teléfono** de un INICIO Entrante | Cruza el número contra `ContactLog.xml` (±120 s) → muestra **ID interno** y **Método de Contestación** (Auto-Answer/manual) desde `$Script:ModoContestacion` |
| Columna **Sesión** | Abre sub-grid filtrando solo las filas de esa sesión (aislar llamada) |

---

## 7. Módulos auxiliares

| Módulo | Mecánica interna |
|--------|------------------|
| **INFO PC** | `Get-WmiObject` (Win32_Processor, _PhysicalMemory, _OperatingSystem, _NetworkAdapter, _ComputerSystem) sobre la IP con credenciales. Calcula RAM, *uptime* desde `LastBootUpTime`. |
| **EXTRAER LOG (RAW)** | Lee Endpoint/OneX/Audio/Ispeac, rastrea hora a nivel **ms**, filtra por rango `HH:mm:ss`, entrelaza por `HoraStr` y colorea por origen. Filtro de texto y export CSV. |
| **BÚSQUEDA EN LOGS** | Regex `(?i)(termino1|termino2…)` sobre todas las fuentes de una fecha. Mantiene flag `$EsFechaCorrecta` para no cruzar días (mismo criterio MM/dd que el motor). Devuelve siempre `@()` para evitar el *unwrap* de PowerShell con 1 resultado. |
| **REINICIOS** | Sube un script `CazaReinicios.ps1` a `\\IP\C$\Users\Public`, crea y ejecuta una **tarea programada** como SYSTEM vía `schtasks`, espera el archivo `.done` (timeout 90 s), lee el resultado y **limpia** tarea + archivos. Analiza eventos `System` 6005/6006/12/13/6008/41 y BSOD `1001`. |
| **EXPORTAR** | Vuelca las columnas **visibles** del grid a CSV UTF-8 (con BOM), escapando comillas. |

---

## 8. Puntos delicados / *gotchas* (leer antes de modificar)

- **Rotación de logs:** `OneXAgent.log` es el segmento actual (puede empezar a
  cualquier hora); `.1`…`.12` son anteriores. **El filtro de fecha DEBE estar
  activo** al leer `$MapeoTelP4`, porque logs históricos repiten los mismos
  `ConnectionId` del día actual → cruces falsos.
- **Slot base "absorbido":** cuando un INICIO saliente se mueve a slot ms, el slot
  base queda con `Interpretacion=""` y marcador `[ABSORBIDO:...]`; el render lo
  oculta. Cuidado al tocar guards de deduplicación (líneas ~1656 y ~2278).
- **PowerShell 5.1:** `Sort-Object` sobre una colección de **un** elemento la
  desenvuelve a string; iterarla con `Where-Object` itera **caracteres**. Fix
  aplicado: envolver con `@()` y castear `[string]$_`.
- **Orden PASO 4 → PASO 5** no es negociable: el FIN definitivo depende de
  EndpointLog.
- **ACD por bandera, no por número:** la clasificación ACD se confirma con
  `Setting Vi.InboundAcd=true`, no asumiendo por el número marcado.
- **Fechas:** nunca aceptar `dd/MM` en los regex de filtro (colisión de meses).

---

## 9. Mapa rápido del código (`AnalizadorOneXV22.ps1`)

| Líneas (aprox.) | Sección |
|-----------------|---------|
| 1–110 | GUI: formulario, controles, grid (10 columnas) |
| 112–173 | Menú contextual + `Add_MouseDown` (habilitar diagnósticos) |
| 178–190 | `Reset-Entorno` |
| 208–240 | **INFO PC** (WMI) |
| 245–289 | **BUSCAR USUARIOS** (ping + PSDrive + credenciales) |
| 294–322 | **RUTA MANUAL** |
| 327–2511 | **ANALIZAR** — motor completo (PASOS 1–6 + barridos) |
| 2514–2742 | Diagnóstico: ¿por qué no hay INICIO? |
| 2746–3128 | Diagnóstico: ¿por qué no hay FIN? |
| 3132–3402 | Diagnóstico: ¿por qué finalizó? |
| 3406–3683 | Diagnóstico: ¿entrante o saliente? |
| 3687–3961 | Detalle: conferencia |
| 3965–4196 | Detalle: transferencia |
| 4200–4232 | **EXPORTAR** CSV |
| 4237–4308 | Clic en celda (evidencia / Auto-Answer / aislar sesión) |
| 4313–4404 | **EXTRAER LOG** (RAW al ms) |
| 4409–4490 | **BÚSQUEDA EN LOGS** |
| 4495–4599 | **REINICIOS** (schtasks remoto) |

> Las líneas son orientativas; el archivo evoluciona. Para ubicar una sección,
> busca el banner `# ====` o el texto `PASO N/6`.

---

*Complemento operativo: [`MANUAL_USUARIO.md`](MANUAL_USUARIO.md).*
