# Manual de Usuario — Analizador Avaya OneX Agent (V22)

> Herramienta de escritorio (PowerShell + WinForms) para auditar la sesión de un
> agente de contact center a partir de los logs del cliente **Avaya OneX Agent**.
> Reconstruye, minuto a minuto, qué pasó con las llamadas, el estado del asesor,
> el audio, la calidad de red y los reinicios del equipo — sin tener que leer los
> archivos de log a mano.
>
> *Actualizado: 08/09/2026.*

---

## 1. ¿Para qué sirve?

Cuando un asesor reporta un problema ("se me cayó la llamada", "no me entran
llamadas", "no pude transferir", "se reinició la máquina", "me puso en un
auxiliar que no elegí"), esta herramienta te arma una **línea de tiempo
legible** de lo que ocurrió en su equipo:

- A qué hora **entró o salió** cada llamada, con qué número, y **cuánto duró**
  la conversación (`mm:ss`).
- Si la llamada fue **entrante (ACD / interna / externa)** o **saliente**.
- Cuándo el asesor se puso **Disponible / Auxiliar** (y con qué motivo, y por
  qué botón), puso la llamada en **espera (Hold)** o activó **Mute**.
- Cómo **terminó** cada llamada (colgó el asesor, colgó el cliente, colgó en
  hold, evasión, falla del motor de auto-contestación, etc.).
- Si hubo **transferencias** o **conferencias**, y cómo se establecieron.
- La **calidad de red/voz** durante la llamada (opcional).
- **Huecos de silencio** en el proceso del cliente OneX (posibles
  congelamientos), aunque no haya ninguna otra señal de que algo falló.
- El **historial de reinicios** del equipo y la **información de hardware**.

Todo se muestra en una sola tabla con colores, y puedes hacer clic en cualquier
celda para ver la **línea de log original** que respalda esa interpretación.

---

## 2. Requisitos

| Requisito | Detalle |
|-----------|---------|
| Sistema | Windows 10 / 11 |
| PowerShell | 5.1 o superior |
| Acceso de red | Solo si analizas un equipo remoto por IP |
| Credenciales | Usuario con acceso administrativo al recurso `\\IP\c$` del equipo remoto |
| Logs | Carpeta `Log Files` de Avaya OneX Agent (ver abajo qué archivos usa) |

**Archivos de log que la herramienta consume:**

- `ContactLog.xml` — fuente principal de llamadas y números.
- `OneXAgent.log` (y `.1`, `.2`, …) — eventos del cliente OneX (incluye estados
  de Disponible/Auxiliar y motivos).
- `EndpointLog.txt` (y `.1`, `.2`, …) — sesiones, hold, fin de llamada, firma.
- `AudioLog.txt` — dispositivos y eventos de audio.
- `IspeacLog.txt` — calidad de red / voz (opcional).

---

## 3. Cómo abrir la herramienta

1. Abre PowerShell.
2. Ejecuta el script:
   ```powershell
   .\AnalizadorOneXV22.ps1
   ```
3. Se abre la ventana **"Analizador OneX V22"**.

> Si tu Windows bloquea la ejecución de scripts, ábrelo con:
> `powershell -ExecutionPolicy Bypass -File .\AnalizadorOneXV22.ps1`

---

## 4. Las dos formas de cargar logs

### Opción A — Equipo remoto (por IP)
Úsala cuando el equipo del asesor está prendido y en red.

1. Escribe la **IP del equipo** en el campo *IP del Equipo*.
2. Clic en **1. BUSCAR USUARIOS**.
   - La herramienta hace un *ping* primero. Si no responde, te avisa (equipo
     apagado o fuera de la red).
   - Te pide **credenciales** (usuario tipo `local\soporte`). Si fallan, te las
     vuelve a pedir automáticamente.
   - Si conecta, llena la lista de **Usuarios Windows** del equipo.
3. Elige el **Usuario Windows** y la **Fecha** del incidente.
4. Clic en **2. ANALIZAR**.

### Opción B — Ruta manual (logs locales)
Úsala cuando ya tienes los logs copiados en tu equipo (o el remoto está apagado).

1. Clic en **3. RUTA MANUAL**.
2. Selecciona la carpeta que contiene los logs (la carpeta `Log Files`).
   - La herramienta detecta automáticamente las fechas disponibles y selecciona
     la más reciente; igual puedes cambiar la fecha libremente.
3. Clic en **2. ANALIZAR**.

---

## 5. Los botones de la ventana

| Botón | Qué hace |
|-------|----------|
| **1. BUSCAR USUARIOS** | Conecta al equipo por IP y lista los usuarios de Windows |
| **2. ANALIZAR** | Procesa los logs y construye la línea de tiempo (el corazón de la herramienta) |
| **3. RUTA MANUAL** | Carga logs desde una carpeta local, sin conexión remota |
| **4. INFO PC** *(aparece tras conectar por IP)* | Muestra hardware del equipo: CPU, RAM, S.O., velocidad de red, *uptime* y usuario en sesión |
| **5. EXTRAER LOG** | Abre la vista **RAW**: las líneas crudas de todos los logs entrelazadas al milisegundo dentro de un rango de horas |
| **6. BÚSQUEDA EN LOGS** | Busca palabras libres (separadas por coma) en todos los logs — de una fecha o de todos los días disponibles |
| **7. REINICIOS** *(equipo remoto)* | Historial de encendidos/apagados/fallas del equipo en los últimos N días |
| **8. EXPORTAR** | Guarda la tabla actual a un archivo **CSV** |
| **🔍 Vacíos en logs (N)** *(aparece solo si hay hallazgos, tras ANALIZAR)* | Lista los huecos de silencio ≥3s detectados en el proceso del cliente OneX durante el día completo — ver [sección 8](#vacíos-en-logs) |
| ☑ **Ver Calidad de Red** | Muestra/oculta la columna de calidad de voz, adjunta a los eventos que ya tienen actividad durante una llamada |

> Los botones se habilitan en orden: primero conecta o carga ruta, luego analiza,
> y después se activan Extraer / Búsqueda / Exportar / Vacíos.

---

## 6. Cómo leer la tabla de resultados

Cada fila es un **momento** de la sesión. Columnas:

| Columna | Qué contiene |
|---------|--------------|
| **Hora** | Hora del evento (HH:mm:ss, a veces con milisegundos) |
| **Sesión** | ID de la llamada/sesión (para agrupar eventos de una misma llamada) |
| **Teléfono** | Número involucrado |
| **Actividad del Agente** | La **interpretación en lenguaje claro** de lo que pasó |
| **Endpoint.log** | Evidencia del archivo EndpointLog |
| **Audio.log** | Evento de audio |
| **AvayaOneX.log** | Evento del cliente OneX |
| **Calidad Red/Voz** | Métricas de calidad (oculta por defecto) |
| **Log de Sistema** | Evento del Visor de Eventos de Windows (System) |
| **Log de Aplicación** | Evento del Visor de Eventos de Windows (Application) |

### Símbolos que verás en "Actividad del Agente"
- ▶ / ► — inicio de llamada
- ▲ — línea abierta / saliente
- ■ — fin de llamada
- → — transferencia / conferencia / actualización de número
- ♪ — música/audio, o calidad de red sana

### Frases típicas y qué significan

| Lo que dice la tabla | Qué significa |
|----------------------|---------------|
| **Agente con señal de llamada (Entrante - ACD / INTERNA / EXTERNA - número)** | Está timbrando una llamada que entra; indica el tipo y el número |
| **INICIO DE LLAMADA (Entrante)** *(con Ring Time)* | El asesor **contestó** una llamada entrante; el *Ring Time* es lo que tardó en contestar |
| **INICIO DE LLAMADA (Saliente)** | El asesor **marcó** una llamada |
| **LÍNEA ABIERTA SIN MARCAR** | Levantó la línea pero no marcó número |
| **INICIO DE SESIÓN (Llamada Interna — Extensión N)** | Llamada interna entre extensiones |
| **Extensión en línea y conectada a dispositivos de audio** | El cliente arrancó y tomó audio correctamente |
| **Usuario intentando firmarse en la Ext. N** | Intento de login del agente |
| **HOLD MANUAL (Clic del Agente)** | El asesor puso la llamada en espera él mismo |
| **HOLD MANUAL — ¡FALLÓ! (no se pudo retener la llamada)** | El asesor intentó poner en hold, pero el sistema no lo logró completar |
| **Hold automático del sistema** | Espera generada por el sistema (p. ej. al abrir otra línea), no por el asesor |
| **Mute activado / desactivado** | Silenció / reactivó su micrófono |
| **CONSULTA DE CONFERENCIA / DE TRANSFERENCIA → número** | Llamó a un tercero para consultar antes de conferenciar o transferir |
| **CONFERENCIA ESTABLECIDA** | Se armó una conferencia (clic derecho → *Ver detalle* para ver cómo) |
| **TRANSFERENCIA COMPLETADA** | Se completó una transferencia (clic derecho → *Ver detalle*) |
| **FIN DE LLAMADA NORMAL** *(con "(habló mm:ss)")* | La llamada terminó de forma normal; el tiempo entre paréntesis es lo que duró la conversación |
| **FIN DE LLAMADA MANUAL (Colgada por el Asesor)** | El asesor colgó |
| **FIN DE LLAMADA - CLIENTE COLGÓ EN HOLD** | El cliente colgó mientras la llamada estaba en espera |
| **CUELGUE MANUAL (Línea abierta sin marcar)** | Colgó una línea que había abierto sin marcar |
| **¡EVASIÓN! Línea abandonada sin marcar (Timeout/Tapón)** | Línea abierta y abandonada sin marcar — posible evasión |
| **⚠ SIN BOTONES PARA TOMAR LA LLAMADA** | La llamada quedó timbrando sin que el asesor pudiera contestarla — falla del motor de auto-contestación |
| **⚠ FALLA EN LA RECEPCIÓN DE LLAMADA** | El motor de auto-contestación del cliente truena (excepción) |
| **Número de cliente actualizado (Transferencia inter-agente)** | El número cambió porque la llamada pasó de un agente a otro |
| **[!] Sesión bridge del sistema (auto-transferencia)** | Sesión técnica creada por el sistema, no una acción del asesor |
| **[CM Auto-Answer]** | La llamada se contestó automáticamente (configuración Auto-Answer del Communication Manager) |

### Estados de Disponible / Auxiliar (motivo)

El asesor puede cambiar su estado de dos formas, y la tabla distingue cuál usó:

- **Botón correcto (esquina superior izquierda de OneX Agent)**: abre el menú
  con los motivos (Comida, Baño, Llamada de salida, Capacitación, Sistemas…).
  Siempre queda un registro **con el motivo exacto y confirmado**:
  - `Asesor se cambia a Disponible (Confirmado por clic)`
  - `Asesor se cambia a Auxiliar [MOTIVO] (Confirmado por clic)`
  - `Asesor se cambia a Default (Confirmado por clic)` — si elige explícitamente "Default".
- **Botón favorito "TrabAux"** (o "Auto-In" para volver a Disponible): son
  atajos rápidos que **no pasan por el menú de motivos**. Aunque el asesor
  teclee un número después de presionar TrabAux (pensando que así elige un
  motivo), **ese número no se envía como código de razón** — no existe esa
  combinación en OneX Agent. Cuando esto pasa, Avaya simplemente **reutiliza
  el último motivo que el asesor eligió explícitamente por el botón
  correcto** en esa misma sesión (o lo deja en **Default** si todavía no ha
  elegido ninguno):
  - `Asesor se cambia a Disponible usando botón favorito AUTO-IN`
  - `Asesor se cambia a Auxiliar [MOTIVO] (vía botón favorito)` — heredó el último motivo real.
  - `Asesor se cambia a Auxiliar - se detectó un auxiliar sin código` — todavía no había ningún motivo elegido esa sesión (queda en Default).
- **Auxiliar elegido con una llamada ACTIVA** (por cualquiera de los dos
  botones): Avaya no lo aplica de inmediato, lo deja **pendiente** hasta que
  cuelga:
  - `... (clic registrado — pendiente, la llamada seguía activa)` — al momento del clic.
  - `Auxiliar pendiente aplicado [MOTIVO]` — cuando finalmente se aplica, al colgar.

> Si un asesor insiste en que "no eligió el auxiliar en el que quedó", revisa
> si usó **TrabAux** en vez del botón de la esquina superior — es la causa más
> común: el sistema no le dio el motivo que esperaba porque TrabAux nunca
> manda ninguno, solo repite el último real.

---

## 7. Acciones extra dentro de la tabla

### Clic en una celda → ver la evidencia
Haz **clic en cualquier celda** (Actividad, Endpoint.log, Audio.log, etc.) y se
abre una ventana con la **línea de log original** que generó esa interpretación.
Es tu respaldo cuando necesitas demostrar lo que pasó.

### Clic en la columna "Teléfono" de un INICIO Entrante → validar Auto-Answer
Si haces clic en el número de un **INICIO DE LLAMADA (Entrante)**, la herramienta
cruza ese número contra `ContactLog.xml` y te dice el **método de contestación**
(Auto-Answer o manual) y el ID interno de la llamada.

### Clic en la columna "Sesión" → aislar la llamada
Haz clic en un **ID de Sesión** y se abre una ventana solo con las filas de **esa
llamada**, para seguirla de principio a fin sin distracciones.

### Clic derecho → menú de diagnóstico
Según la fila, el menú contextual ofrece:

- **¿Por qué no hay INICIO DE LLAMADA?** — sobre una fila de "señal de llamada".
- **¿Por qué no hay FIN DE LLAMADA?** — sobre un INICIO / LÍNEA ABIERTA.
- **¿Por qué finalizó esta llamada?** — sobre un FIN / CUELGUE.
- **¿Por qué es Entrante o Saliente?** — sobre un INICIO DE LLAMADA.
- **Ver detalle: ¿Cómo se estableció la Conferencia?**
- **Ver detalle: ¿Cómo se completó la Transferencia?**

Estas opciones abren un reporte paso a paso que explica el razonamiento de la
herramienta (muy útil cuando un caso se ve raro).

---

## 8. Herramientas complementarias

### EXTRAER LOG (vista RAW al milisegundo)
1. Clic en **5. EXTRAER LOG**.
2. Indica **Hora Inicio** y **Hora Fin** (formato `HH:mm:ss`).
3. Se abre una vista con **todas las líneas crudas** de Endpoint, OneX, Audio e
   Ispeac, **entrelazadas por hora** y coloreadas por origen. Puedes **filtrar**
   por texto y **exportar a CSV**.

Úsala cuando necesitas el detalle fino, segundo a segundo, de un tramo concreto.

### BÚSQUEDA EN LOGS
1. Clic en **6. BÚSQUEDA EN LOGS**.
2. Escribe palabras separadas por coma (ej. `LogoutRequest, AgentState`).
3. Elige la fecha, o marca ☑ **"Buscar en todos los días disponibles"** si no
   sabes en qué fecha ocurrió lo que buscas (por ejemplo, un número o sesión
   que sospechas que viene de un día distinto al que estás analizando).
4. Clic en **BUSCAR**. Si buscaste en todos los días, el resultado incluye una
   columna **Fecha** y queda ordenado cronológicamente.
5. **EXPORTAR CSV** guarda los resultados de la búsqueda actual a un archivo.

### 🔍 Vacíos en logs
Tras **2. ANALIZAR**, si el proceso del cliente OneX estuvo **≥3 segundos sin
generar ninguna línea de log** en algún momento del día (posible
congelamiento), aparece el botón **"🔍 Vacíos en logs (N)"** con el número de
huecos encontrados. Al abrirlo, verás cada hueco con su hora de inicio, hora en
que se reanuda, duración, y si **había una llamada activa** en ese momento
(corroborado contra `IspeacLog`, que corre en un proceso aparte) — un hueco con
llamada activa es más grave, porque significa que el cliente se congeló *con
el cliente en línea*.

Útil para casos donde el asesor reporta "se trabó" sin que el timeline muestre
nada obvio — el vacío en sí es la evidencia.

### REINICIOS
Sobre un equipo remoto, consulta los últimos *N* días y clasifica cada evento:
- **[OK]** reinicio limpio,
- **[ALERTA]** equipo encendido muchos días seguidos,
- **[PELIGRO]** apagón / botonazo / pantalla azul,
- **[ACTUAL]** tiempo encendido actual.

### INFO PC
Reporte rápido de hardware del equipo remoto (CPU, RAM, S.O., red, *uptime*,
usuario en sesión).

---

## 9. Flujo recomendado para un caso típico

1. Recibe el reporte del asesor (hora aproximada y síntoma).
2. Conecta por **IP** (o carga **Ruta Manual** si te pasaron los logs).
3. Elige usuario y la **fecha** del incidente → **ANALIZAR**.
4. Ubica la hora en la tabla y lee la **Actividad del Agente**.
5. Si algo no cuadra, **clic derecho → Diagnosticar** para ver el razonamiento.
6. Para el detalle fino, usa **EXTRAER LOG** en ese rango de horas.
7. Si sospechas que algo viene de otro día (número raro, sesión reciclada),
   usa **BÚSQUEDA EN LOGS** con la casilla "todos los días" activada.
8. Si el asesor reporta que "se trabó" sin causa clara, revisa **🔍 Vacíos en
   logs**.
9. **Exporta a CSV** para adjuntar la evidencia al ticket.

---

## 10. Problemas frecuentes

| Síntoma | Causa probable / solución |
|---------|---------------------------|
| "El equipo no responde al ping" | Equipo apagado o fuera de red → usa **Ruta Manual** con los logs copiados |
| Te pide credenciales una y otra vez | Usuario/contraseña incorrectos o sin permiso a `\\IP\c$` |
| La tabla sale vacía | Fecha equivocada, o esa carpeta no tiene logs de ese día — si no sabes la fecha correcta, usa **BÚSQUEDA EN LOGS** con "todos los días" |
| No veo la columna de calidad | Activa la casilla **Ver Calidad de Red** |
| "Primero debes realizar una auditoría" | Los botones 5/6 requieren haber corrido **ANALIZAR** antes |
| No aparece **INFO PC** / **REINICIOS** | Solo se muestran tras conectar por **IP** (no en modo Ruta Manual) |
| El asesor dice que no eligió el auxiliar en que quedó | Revisa si usó el botón favorito **TrabAux** — no permite elegir motivo, hereda el último real (ver sección 6) |

---

*Para entender **cómo** la herramienta llega a cada interpretación (la mecánica
interna, el pipeline y la lógica de detección), consulta el documento
[`MANUAL_TECNICO_SOPORTE.md`](MANUAL_TECNICO_SOPORTE.md).*
