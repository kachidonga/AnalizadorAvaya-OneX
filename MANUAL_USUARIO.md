# Manual de Usuario — Analizador Avaya OneX Agent (V22)

> Herramienta de escritorio (PowerShell + WinForms) para auditar la sesión de un
> agente de contact center a partir de los logs del cliente **Avaya OneX Agent**.
> Reconstruye, minuto a minuto, qué pasó con las llamadas, el estado del asesor,
> el audio, la calidad de red y los reinicios del equipo — sin tener que leer los
> archivos de log a mano.

---

## 1. ¿Para qué sirve?

Cuando un asesor reporta un problema ("se me cayó la llamada", "no me entran
llamadas", "no pude transferir", "se reinició la máquina"), esta herramienta te
arma una **línea de tiempo legible** de lo que ocurrió en su equipo:

- A qué hora **entró o salió** cada llamada, y con qué número.
- Si la llamada fue **entrante (ACD / interna / externa)** o **saliente**.
- Cuándo el asesor se puso **Disponible / No disponible**, puso en **espera (Hold)**
  o activó **Mute**.
- Cómo **terminó** cada llamada (colgó el asesor, colgó el cliente, evasión, etc.).
- Si hubo **transferencias** o **conferencias**, y cómo se establecieron.
- La **calidad de red/voz** durante la llamada (opcional).
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
- `OneXAgent.log` (y `.1`, `.2`, …) — eventos del cliente OneX.
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
| **6. BÚSQUEDA EN LOGS** | Busca palabras libres (separadas por coma) en todos los logs de una fecha |
| **7. REINICIOS** *(equipo remoto)* | Historial de encendidos/apagados/fallas del equipo en los últimos N días |
| **8. EXPORTAR** | Guarda la tabla actual a un archivo **CSV** |
| ☑ **Ver Calidad de Red** | Muestra/oculta la columna de calidad de voz y las filas "Llamada en curso (Analizando Calidad)" |

> Los botones se habilitan en orden: primero conecta o carga ruta, luego analiza,
> y después se activan Extraer / Búsqueda / Exportar.

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
- ♪ — música/audio

### Frases típicas y qué significan

| Lo que dice la tabla | Qué significa |
|----------------------|---------------|
| **Agente con señal de llamada (Entrante - ACD / INTERNA / EXTERNA - número)** | Está timbrando una llamada que entra; indica el tipo y el número |
| **INICIO DE LLAMADA (Entrante)** *(con Ring Time)* | El asesor **contestó** una llamada entrante; el *Ring Time* es lo que tardó en contestar |
| **INICIO DE LLAMADA (Saliente)** | El asesor **marcó** una llamada |
| **LÍNEA ABIERTA SIN MARCAR** | Levantó la línea pero no marcó número |
| **INICIO DE SESIÓN (Llamada Interna — Extensión N)** | Llamada interna entre extensiones |
| **Asesor se cambia a Disponible** *(Auto-In / Confirmado por clic)* | Se puso en estado disponible para recibir llamadas |
| **Extensión en línea y conectada a dispositivos de audio** | El cliente arrancó y tomó audio correctamente |
| **Usuario intentando firmarse en la Ext. N** | Intento de login del agente |
| **HOLD MANUAL (Clic del Agente)** | El asesor puso la llamada en espera él mismo |
| **Hold automático del sistema** | Espera generada por el sistema (p. ej. al abrir otra línea), no por el asesor |
| **Mute activado / desactivado** | Silenció / reactivó su micrófono |
| **CONSULTA DE CONFERENCIA / DE TRANSFERENCIA → número** | Llamó a un tercero para consultar antes de conferenciar o transferir |
| **CONFERENCIA ESTABLECIDA** | Se armó una conferencia (clic derecho → *Ver detalle* para ver cómo) |
| **TRANSFERENCIA COMPLETADA** | Se completó una transferencia (clic derecho → *Ver detalle*) |
| **FIN DE LLAMADA NORMAL** | La llamada terminó de forma normal |
| **FIN DE LLAMADA MANUAL (Colgada por el Asesor)** | El asesor colgó |
| **CUELGUE MANUAL (Línea abierta sin marcar)** | Colgó una línea que había abierto sin marcar |
| **¡EVASIÓN! Línea abandonada sin marcar (Timeout/Tapón)** | Línea abierta y abandonada sin marcar — posible evasión |
| **Número de cliente actualizado (Transferencia inter-agente)** | El número cambió porque la llamada pasó de un agente a otro |
| **[!] Sesión bridge del sistema (auto-transferencia)** | Sesión técnica creada por el sistema, no una acción del asesor |
| **[CM Auto-Answer]** | La llamada se contestó automáticamente (configuración Auto-Answer del Communication Manager) |

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
3. Elige la fecha y clic en **BUSCAR**.
4. Te muestra cada coincidencia con su hora y archivo de origen.

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
7. **Exporta a CSV** para adjuntar la evidencia al ticket.

---

## 10. Problemas frecuentes

| Síntoma | Causa probable / solución |
|---------|---------------------------|
| "El equipo no responde al ping" | Equipo apagado o fuera de red → usa **Ruta Manual** con los logs copiados |
| Te pide credenciales una y otra vez | Usuario/contraseña incorrectos o sin permiso a `\\IP\c$` |
| La tabla sale vacía | Fecha equivocada, o esa carpeta no tiene logs de ese día |
| No veo la columna de calidad | Activa la casilla **Ver Calidad de Red** |
| "Primero debes realizar una auditoría" | Los botones 5/6 requieren haber corrido **ANALIZAR** antes |
| No aparece **INFO PC** / **REINICIOS** | Solo se muestran tras conectar por **IP** (no en modo Ruta Manual) |

---

*Para entender **cómo** la herramienta llega a cada interpretación (la mecánica
interna, el pipeline y la lógica de detección), consulta el documento
[`MANUAL_TECNICO_SOPORTE.md`](MANUAL_TECNICO_SOPORTE.md).*
