[English (EN)](README.md) 🇬🇧 | [Español (ES)](README.es.md) 🇪🇸

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
- **Integración con Windows Property System**: Extrae metadatos de las cabeceras FITS y los inyecta de forma nativa en el Explorador de Windows. *Nota: Solo se indexan los metadatos concretos enumerados abajo, ya que el código C# debe mapearlos contra identificadores estándar de Windows (`System.*`). No es posible indexar variables arbitrarias de FITS simplemente modificando el registro, requeriría cambiar el código fuente.*
  - **Metadatos Mapeados e Indexados**:
    - `System.Subject` (mapeado desde el FITS `OBJECT`)
    - `System.Image.HorizontalSize` (desde `NAXIS1`)
    - `System.Image.VerticalSize` (desde `NAXIS2`)
    - `System.Image.BitDepth` (desde `BITPIX`)
    - `System.Photo.CameraModel` (desde `INSTRUME` o `CAMERA`)
    - `System.Photo.ExposureTime` (desde `EXPOSURE` o `EXPTIME`)
  - **Visualización en el Sistema**:
    - **InfoTip Nativo**: Mantén el ratón sobre un archivo FITS para ver una tarjeta emergente con el resumen de la captura.
    - **Pestaña Detalles**: Pulsa `Alt+Enter` (o abre Propiedades -> Detalles) para ver los metadatos extraídos. El script de registro configura nativamente la disposición de estos datos (`FullDetails`, `InfoTip`, `PreviewDetails`).
  - **Búsqueda Avanzada en el Explorador de Windows**:
    Al estar indexados, Windows permite usar su Sintaxis de Consulta Avanzada (AQS) directamente en la barra de búsqueda superior derecha del Explorador. Como los nombres localizados ("cámara", "exposición") pueden variar, **es recomendable usar el nombre del sistema (`System.*`)**:
    - **Búsqueda exacta o texto**: `System.Photo.CameraModel:ZWO` o `System.Subject:M31`
    - **Filtros numéricos y rangos**: Filtra tiempos de exposición usando matemáticas. Ejemplo: `System.Photo.ExposureTime:>120` (más de 120s), o un rango `System.Photo.ExposureTime:10..300` (entre 10 y 300 segundos).
    - **Combinando Condiciones (AND / OR)**: Usa lógica booleana encadenada (siempre en MAYÚSCULAS). 
      *Ejemplo: `System.Photo.CameraModel:"QHY" AND System.Photo.ExposureTime:>120` (Cámara QHY y tomas de más de 2 minutos).*
- **Layout Dinámico**: Imagen arriba y tabla de datos abajo con altura redimensionable.
- **Streaming Progresivo**: Muestra el progreso de lectura centrado sobre la imagen mientras se procesa el flujo de datos.
- **Menú Contextual Integrado**: Haz clic derecho en cualquier lugar de la vista previa para configurar y exportar:
  - **Mostrar / Ocultar Imagen**: Cambia instantáneamente entre renderizar la imagen FITS o cargar solo la matriz de metadatos ultra-rápida.
  - **Activar / Desactivar Trazas**: Habilita los logs de diagnóstico para depuración en un clic.
  - **Copiar Imagen**: Copia la vista auto-estirada de la imagen al portapapeles (ideal para pegar en Paint, Photoshop o mensajes).
  - **Copiar Fila Seleccionada**: Copia exactamente la fila de metadatos seleccionada al portapapeles.
  - **Copiar Toda la Tabla (CSV)**: Exporta todo el array del Header FITS en formato CSV (con cabeceras) al portapapeles.

## ⚙️ Configuración (Registro)

La configuración se gestiona fácilmente desde el propio **Menú Contextual (Clic derecho)** en la ventana de previsualización, sin necesidad de tocar `regedit.exe`. Los cambios surten efecto al instante.

Para automatizaciones o despliegues, las claves de registro se almacenan de forma segura en la zona de Baja Integridad (Low-IL):

### Nivel de Usuario (Recomendado)
`HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`

### Nivel Global (Predeterminado)
`HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Valor (DWORD) | Descripción |
| :--- | :--- |
| `ShowImage` | `1` (Mostrar imagen y tabla), `0` (Solo tabla). |
| `EnableTracing` | `1` (Habilitar logs de depuración), `0` (Deshabilitar). |

---

## 🚀 Instalación (Como Administrador)

1.  Compila el proyecto con `dotnet build -c Release`.
2.  Ejecuta `register.bat` para dar de alta el componente en el sistema.
3.  Si necesitas desinstalar, usa `unregister.bat`.

---

## 🔍 Diagnóstico (Logs)

Los logs se guardan en la zona de baja integridad para cumplir con las restricciones de `prevhost.exe`:
`%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`
