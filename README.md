# Analizador Avaya OneX Agent

Herramienta de escritorio en PowerShell para auditar logs del cliente **Avaya OneX Agent**. Permite analizar la sesión de un agente de contact center — eventos de llamada, calidad de red, reinicios y más — sin necesidad de acceder manualmente a los archivos de log.

## Funcionalidades principales

| Botón | Acción |
|-------|--------|
| 1. BUSCAR USUARIOS | Conecta al equipo por IP y lista los usuarios de Windows disponibles |
| 2. AUDITAR | Analiza los logs del usuario y fecha seleccionados |
| 3. RUTA MANUAL | Carga logs desde una carpeta local (sin conexión remota) |
| 5. EXTRAER LOG | Copia los archivos de log del equipo remoto al equipo local |
| 6. BÚSQUEDA | Busca texto libre dentro de los eventos cargados |
| 7. REINICIOS | Detecta reinicios de la aplicación en la sesión analizada |

## Requisitos

- Windows 10 / 11
- PowerShell 5.1 o superior
- Acceso de red al equipo remoto (para análisis por IP)
- Credenciales de administrador del dominio (para montar la unidad de red)

## Cómo usar

1. Abre PowerShell y ejecuta el script:
   ```powershell
   .\AnalizadorOneXV21-1.ps1
   ```
2. Ingresa la IP del equipo a analizar y haz clic en **1. BUSCAR USUARIOS**.
3. Selecciona el usuario de Windows y la fecha del incidente.
4. Haz clic en **2. AUDITAR** para cargar y analizar los eventos.
5. Usa los botones adicionales para extraer logs, buscar eventos o ver reinicios.

> Si los logs ya están en tu equipo local, usa **3. RUTA MANUAL** para omitir la conexión remota.

## Estructura del proyecto

```
OneX/
├── AnalizadorOneXV21-1.ps1   # Versión estable más reciente
├── AnalizadorOneXV21.ps1     # Versión anterior
├── AuditoriaV*.ps1           # Versiones históricas del analizador de auditoría
├── Logs de ejemplo/          # Logs de referencia para pruebas
└── cambios*.txt              # Registro de ideas y cambios pendientes
```

## Versiones

- **V21-1** — Versión estable actual
- **V21** — Refactorización de flujo principal
- **V20** — Incorporación de módulo de búsqueda y extracción de logs
