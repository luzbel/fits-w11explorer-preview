[ESPAÑOL](#fits-windows-explorer-preview-handler-español) | [ENGLISH](#fits-windows-explorer-preview-handler-english)

---

# FITS Windows Explorer Preview Handler (ESPAÑOL)

> **Previsualizador nativo de archivos FITS para el Explorador de Windows 10/11.**

![Platform: Windows 10/11](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)
![Zero-Copy](https://img.shields.io/badge/Architecture-Zero--Copy-green)

---

## ✨ Características Principales

- **Arquitectura Zero-Copy**: Lee archivos FITS de varios GB de forma instantánea sin copiarlos a disco, usando flujos directos (`IStream`).
- **Previsualización de Imagen**: Renderiza el contenido del canal principal con soporte para:
  - **Debayering Inteligente**: Downsampling 2x2 (Binning) para sensores color sin artefactos de rejilla.
  - **Auto-Stretch Adaptativo**: Ajuste dinámico de niveles basado en Mediana/MAD para ver objetos débiles.
  - **Tintado por Filtro**: Colorea automáticamente tomas de banda estrecha (Ha, OIII, SII) según el header.
- **Layout Dinámico**: Imagen arriba y tabla de datos abajo con altura redimensionable.
- **Streaming Progresivo**: Muestra el progreso de lectura centrado sobre la imagen mientras se procesa el flujo de datos.
- **Tabla de Metadatos**: Muestra todos los keywords, valores y comentarios del header FITS con resaltado de sintaxis.

## ⚙️ Configuración (Registry)

El comportamiento del previsualizador se puede ajustar mediante el Registro de Windows. Los cambios son instantáneos al seleccionar el siguiente archivo.

### Nivel de Usuario (Recomendado)
Usa esta ruta para configuraciones personales que no requieren permisos de administrador:
`HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`

### Nivel Global (Predeterminado)
Configuración para todos los usuarios del sistema (requiere administrador):
`HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Valor (DWORD) | Descripción |
| :--- | :--- |
| `ShowImage` | `1` (Mostrar imagen y tabla), `0` (Solo tabla). |
| `EnableTracing` | `1` (Habilitar logs de depuración), `0` (Deshabilitar). |

---

## 🚀 Instalación (Como Administrador)

1.  Compila el proyecto con `dotnet build`.
2.  Ejecuta `register.bat` para dar de alta el componente en el sistema.
3.  Si necesitas desinstalar, usa `unregister.bat`.

---

## 🔍 Diagnóstico (Logs)

Los logs se guardan en la zona de baja integridad para cumplir con las restricciones de `prevhost.exe`:
`%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`

---
---

# FITS Windows Explorer Preview Handler (ENGLISH)

> **Native FITS file previewer for Windows 10/11 Explorer.**

![Platform: Windows 10/11](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)
![Zero-Copy](https://img.shields.io/badge/Architecture-Zero--Copy-green)

---

## ✨ Main Features

- **Zero-Copy Architecture**: Reads multi-GB FITS files instantly without copying to disk, using direct streams (`IStream`).
- **Image Preview**: Renders the main channel content with support for:
  - **Smart Debayering**: 2x2 Downsampling (Binning) for color sensors without grid artifacts.
  - **Adaptive Auto-Stretch**: Dynamic level adjustment based on Median/MAD to reveal faint objects.
  - **Filter Tinting**: Automatically colors narrowband shots (Ha, OIII, SII) based on headers.
- **Dynamic Layout**: Image on top and metadata table at the bottom with resizable height.
- **Progressive Streaming**: Displays reading progress over the image area while processing the data stream.
- **Metadata Table**: Shows all keywords, values, and comments from the FITS header with syntax highlighting.

## ⚙️ Configuration (Registry)

The previewer behavior can be adjusted via the Windows Registry. Changes are effective immediately upon selecting the next file.

### User Level (Recommended)
Use this path for personal settings that do not require administrator privileges:
`HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`

### Global Level (Default)
System-wide configuration for all users (requires administrator):
`HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Value (DWORD) | Description |
| :--- | :--- |
| `ShowImage` | `1` (Show image and table), `0` (Table only). |
| `EnableTracing` | `1` (Enable debug logs), `0` (Disable). |

---

## 🚀 Installation (As Administrator)

1.  Build the project using `dotnet build`.
2.  Run `register.bat` to register the component in the system.
3.  If you need to uninstall, use `unregister.bat`.

---

## 🔍 Diagnostics (Logs)

Logs are saved in the Low-Integrity area to comply with `prevhost.exe` restrictions:
`%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`
