[English (EN)](README.md) 🇬🇧 | [Español (ES)](README.es.md) 🇪🇸

# Extensión de Vista Previa FITS para Windows

> **Visor nativo y motor de indexación de archivos FITS para el Explorador de Windows 10/11.**

![Plataforma: Windows 10/11](https://img.shields.io/badge/Plataforma-Windows%2010%2F11-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)
![Responsive](https://img.shields.io/badge/Arquitectura-Responsiva-green)

---

## ✨ Características Principales

- **Arquitectura Responsiva**: Lee archivos FITS de varios GB instantáneamente sin copiarlos a disco.
- **E/S Optimizada para la Nube**: Utiliza una estrategia de **Lectura en Bloque (32MB)** para minimizar el tráfico de red en almacenamientos de alta latencia (pCloud, OneDrive, NAS).
- **Diseño "Instant Release"**: Libera los bloqueos de archivo en milisegundos tras la lectura. Puedes renombrar o mover carpetas incluso mientras el renderizado de alta resolución se procesa en RAM.
- **Navegación sin Esperas**: La descarga asíncrona asegura **0ms de retraso** al desplazarte por cientos de archivos con las flechas del teclado.
- **Visualización Prioritaria de Metadatos**: La tabla de cabeceras FITS aparece al instante, mientras la imagen se carga en segundo plano con información de progreso.
- **Proveedor de Miniaturas (Thumbnails)**: Genera miniaturas nativas para archivos `.fits`:
  - **Muestreo por Zancada (Stride)**: Solo lee los píxeles necesarios para el tamaño del icono.
  - **Integridad Bayer**: Maneja correctamente los patrones Bayer para evitar rayas verticales.
  - **Modo Badge estático**: Si `ShowImage=0`, las miniaturas muestran una ficha coloreada:
    - **Fondo**: Codifica el tipo de frame (`LIGHT`=Azul oscuro, `FLAT`=Gris claro, `DARK`=Rojo oscuro, `BIAS`=Gris oscuro).
    - **Banda Superior**: Codifica el filtro (ej: Rojo para Ha, Cian para OIII).
  - **Detección de Latencia**: Mide el tiempo de respuesta del disco. Si es >200ms (nube sin caché), muestra un **Badge** al instante para evitar que el Explorador se cuelgue.
  - **Actualización Automática**: Al ver un archivo en el Panel de Vista Previa, se fuerza el refresco de su miniatura para pasar de "Badge" a "Foto" una vez hidratado en disco.
  - **Etiquetas**: Texto de alto contraste con el tipo de frame y formato ("FITS").
- **Integración con el Sistema de Propiedades**: Extrae metadatos y los inyecta nativamente en Windows.
  - **Propiedades Mapeadas**:
    - `System.Subject` (desde FITS `OBJECT`)
    - `System.Image.HorizontalSize` (desde `NAXIS1`)
    - `System.Image.VerticalSize` (desde `NAXIS2`)
    - `System.Image.BitDepth` (desde `BITPIX`)
    - `System.Photo.CameraModel` (desde `CAMERA` / `INSTRUME`)
    - `System.Photo.ExposureTime` (desde `EXPTIME` / `EXPOSURE`)
    - `System.Category` (desde `IMAGETYP` / `FRAME` normalizado a `Light`/`Dark`/`Flat`/`Bias`)
  - **Optimización de respuesta**: Omite archivos no válidos instantáneamente (bytes mágicos), asegurando que los Tooltips (InfoTip) sean fluidos.
  - **Búsqueda Avanzada**: Usa AQS en la barra del Explorador para filtrar tu biblioteca:
    | Ejemplo de Búsqueda | Resultado |
    | :--- | :--- |
    | `System.Category:Light` | Encuentra todos los Light frames |
    | `System.Category:Dark` | Encuentra todos los Dark frames |
    | `System.Category:Flat` | Encuentra todos los Flat de calibración |
    | `System.Category:Bias` | Encuentra todos los Bias |
    | `System.Photo.ExposureTime:>300` | Exposiciones de más de 5 minutos |
    | `System.Subject:M31` | Tomas de la Galaxia de Andrómeda |
    | `System.Photo.CameraModel:ASI2600` | Archivos de una cámara específica |
    | **Combinado**: `System.Category:Light AND System.Subject:M31` | Lights de un objeto concreto |
    | **Combinado**: `System.Category:Dark AND System.Photo.ExposureTime:300` | Darks de exactamente 300s |

---

## 🎨 Diseño Visual y Experiencia

El panel de vista previa ha sido rediseñado para máxima productividad:

1.  **Panel Superior (Dinámico)**:
    -   Imagen con estiramiento automático (Adaptive Median/MAD).
    -   Progreso en tiempo real (`Cargando...`, `Leyendo 45%...`) para unidades de red o nube lentas.
2.  **Panel Inferior (Metadatos)**:
    -   Tabla de cabeceras FITS totalmente desplazable.
    -   Keywords coloreadas (Rojo para `END`, Amarillo para `NAXIS`, Verde para comentarios).
3.  **Barra de Estado (Pie de página)**:
    -   **Izquierda**: Ayuda de configuración ("Clic derecho para opciones").
    -   **Right**: Versión de la APP y fecha/hora de compilación.

---

## 🖱️ Menú Contextual (Botón Derecho)

Haz clic derecho en cualquier lugar del panel de vista previa:
-   **Modos de Rendimiento**:
    -   `Mostrar Imagen`: Renderizado de alta calidad y miniaturas reales.
    -   `Ocultar Imagen`: Modo ultra-rápido de solo metadatos (ideal para revisar miles de ficheros).
-   **Portapapeles y Exportación**:
    -   `Copiar Imagen`: Genera un BMP estirado de alta calidad y lo copia al portapapeles.
    -   `Copiar Fila Seleccionada`: Copia la línea `Keyword = Value / Comment` seleccionada.
    -   `Copiar Tabla Completa (CSV)`: Exporta toda la cabecera FITS como un string CSV (con títulos) para Excel o Google Sheets.
-   **Configuración y Diagnósticos**:
    -   `Activar/Desactivar Trazas`: Controla los registros de depuración en la carpeta AppDataLow.

---

## 🛠️ Arquitectura

### El motor SampledStream
- **Bufferizado Selectivo**: Almacena en RAM la cabecera y unas ~600 filas de píxeles muestreadas.
- **Liberación Temprana**: Una vez en RAM, el flujo original se **libera inmediatamente**.
- **Filas Virtuales**: Utiliza repetición por vecino más próximo con paridad para preservar patrones Bayer durante el reescalado de miniaturas.

---

## ⚙️ Configuración (Registro)

La extensión comprueba los ajustes en este orden:
1.  **Usuario Local**: `HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`
2.  **Máquina Global**: `HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Valor (DWORD) | Descripción |
| :--- | :--- |
| `ShowImage` | `1` (Renderizar imagen), `0` (Solo metadatos). |
| `EnableTracing` | `1` (Activa registros en AppDataLow). |

---

## 🚀 Instalación y Diagnósticos
1.  Ejecutar `register.bat` como Administrador.
2.  Logs: `%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`
