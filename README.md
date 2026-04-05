# FITS Windows Explorer Preview Handler

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
- **Layout Dinámico**: Imagen arriba y tabla de datos abajo.
  - **Altura Redimensionable**: Arrastra el divisor horizontal para dar más espacio a la imagen o a la tabla.
  - **Scroll Independiente**: La tabla de metadatos mantiene su propio scrollbar para facilitar la navegación por headers extensos sin perder de vista la imagen.
- **Streaming Progresivo**: Lee archivos de gigas sin bloquear la interfaz. Muestra el progreso de lectura centrado sobre la imagen mientras se procesa el flujo de datos.
- **Tabla de Metadatos**: Muestra todos los keywords, valores y comentarios del header FITS.
  - 🟡 **Amarillo**: `NAXIS*` (Dimensiones).
  - 🟢 **Verde**: `COMMENT` / `HISTORY`.
  - 🔴 **Rojo**: `END`.

## 🛠️ Requisitos

- **Windows 10 / 11** (64-bit).
- **.NET Framework 4.8**.

## 🚀 Compilación e Instalación

### Compilar
```powershell
dotnet build FitsPreviewHandler.csproj -c Debug
```

### Instalar (Como Administrador)
Usa los archivos por lotes incluidos para una gestión sencilla:
1.  Ejecuta `unregister.bat` para limpiar registros anteriores.
2.  Ejecuta `register.bat` para dar de alta el componente.

## 📁 Estructura del Proyecto

- `FitsPreviewHandlerExtension.cs`: Implementación de interfaces COM de Windows.
- `FitsPreviewControl.cs`: UI (WinForms) y motor de procesado FITS.
- `ComStreamWrapper.cs`: Puente de alto rendimiento entre el Shell de Windows y .NET.
- `register.bat` / `unregister.bat`: Scripts de despliegue rápido.

## 🔍 Diagnóstico (Logs)

Debido a las restricciones de seguridad de `prevhost.exe`, los logs y datos temporales se almacenan en la zona de baja integridad del usuario:

```
%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log
```

Puedes monitorizar la carga en tiempo real usando herramientas como **DbgView** de Sysinternals (Capture Win32).

## 🗺️ Hoja de Ruta
Consulta el archivo [ROADMAP.md](ROADMAP.md) para ver las próximas mejoras, incluyendo la **Vista de Detalles** (integración de metadatos en columnas del Explorador).
