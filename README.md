[English (EN)](README.md) 🇬🇧 | [Español (ES)](README.es.md) 🇪🇸

# FITS Windows Explorer Preview Handler

> **Native FITS file previewer and metadata indexer for Windows 10/11 Explorer.**

![Platform: Windows 10/11](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)
![Responsive](https://img.shields.io/badge/Architecture-Responsive-green)

---

## ✨ Main Features

- **Responsive Architecture**: Reads multi-GB FITS files instantly without copying to disk.
- **Near-Zero Lock Design**: Releases file handles in milliseconds after an initial sampling phase. You can rename or move folders even while the preview is active.
- **Metadata-First Display**: The FITS header table appears instantly, while the image loads in the background with progress feedback.
- **Thumbnail Provider**: Generates native Windows Explorer thumbnails for `.fits` files:
  - **Stride Sampling**: Only reads the fraction of pixels needed for the icon size.
  - **Bayer Integrity**: Correctly handles Bayer patterns to avoid vertical "stripe" artifacts.
  - **Static Badge mode**: If `ShowImage=0`, thumbnails show a colour-coded card:
    - **Background**: Encodes frame type (`LIGHT`=Dark Blue, `FLAT`=Light Grey, `DARK`=Dark Red, `BIAS`=Dark Grey).
    - **Top Stripe**: Encodes the filter (e.g., Red for Ha, Cyan for OIII, etc.).
    - **Labels**: High-contrast text identifying the frame type and format ("FITS").
- **Windows Property System Integration**: Extracts metadata from FITS headers and injects them natively into Windows Explorer.
  - **Mapped Properties**:
    - `System.Subject` (from FITS `OBJECT`)
    - `System.Image.HorizontalSize` (from `NAXIS1`)
    - `System.Image.VerticalSize` (from `NAXIS2`)
    - `System.Image.BitDepth` (from `BITPIX`)
    - `System.Photo.CameraModel` (from `CAMERA` / `INSTRUME`)
    - `System.Photo.ExposureTime` (from `EXPTIME` / `EXPOSURE`)
    - `System.Category` (from `IMAGETYP` / `FRAME` normalized to `Light`/`Dark`/`Flat`/`Bias`)
  - **Fast-Bail Optimization**: Immediately skips non-FITS files by checking magic bytes, ensuring Tooltips (InfoTip) stay fast.
  - **Advanced Search**: Use Windows AQS in the Explorer search bar to filter your library:
    | Query Example | Result |
    | :--- | :--- |
    | `System.Category:Light` | Finds all Light frames |
    | `System.Category:Dark` | Finds all Dark frames |
    | `System.Category:Flat` | Finds all Flat calibration frames |
    | `System.Category:Bias` | Finds all Bias frames |
    | `System.Photo.ExposureTime:>300` | Finds exposures longer than 5 minutes |
    | `System.Subject:M31` | Finds all shots of the Andromeda Galaxy |
    | `System.Photo.CameraModel:ASI2600` | Finds files from a specific camera |
    | **Combined**: `System.Category:Light AND System.Subject:M31` | Find lights of a specific target |
    | **Combined**: `System.Category:Dark AND System.Photo.ExposureTime:300` | Find darks of exactly 300s |

---

## 🎨 Visual Layout & Experience

The preview pane has been redesigned for maximum productivity:

1.  **Top Panel (Dynamic)**:
    -   Auto-stretched image preview (Adaptive Median/MAD).
    -   Real-time progress (`Loading...`, `Reading 45%...`) for slow network/Cloud drives.
2.  **Bottom Panel (Metadata)**:
    -   Fully scrollable/selectable grid.
    -   Colour-coded keywords (Red for `END`, Yellow for `NAXIS`, Green for comments).
3.  **Status Bar (Footer)**:
    -   **Left**: Configuration hint ("Right-click for options").
    -   **Right**: Application version and Compilation timestamp.

---

## 🖱️ Context Menu (Right-Click)

Right-click anywhere inside the preview pane:
-   **Performance Modes**:
    -   `Show Image`: Renders the high-quality auto-stretched image and real thumbnails.
    -   `Hide Image`: Metadata-only mode. Ultra-fast for browsing thousands of files.
-   **Clipboard & Export**:
    -   `Copy Image`: Generates a high-quality stretched BMP and copies it to your clipboard.
    -   `Copy Selected Row`: Copies the selected `Keyword = Value / Comment` line.
    -   `Copy Full Table (CSV)`: Exports the entire FITS header as a CSV string (with headers) to Excel/Google Sheets.
-   **Configuration & Diagnostics**:
    -   `Enable/Disable Trace`: Toggles debug logs in the AppDataLow folder.

---

## 🛠️ Architecture

### The SampledStream Engine
- **Selective Memory Buffering**: Buffers header and ~600 sampled pixel rows.
- **Early Release**: Once buffered, the original file handle is **immediately released**.
- **Virtual Rows**: Uses Parity-Aware Nearest Neighbor logic to preserve Bayer patterns during thumbnail resizing.

---

## ⚙️ Configuration (Registry)

The extension checks settings in this order:
1.  **Local User**: `HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`
2.  **Global Machine**: `HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Value (DWORD) | Description |
| :--- | :--- |
| `ShowImage` | `1` (Render image), `0` (Metadata only). |
| `EnableTracing` | `1` (Enable logs in AppDataLow). |

---

## 🚀 Installation & Diagnostics
1.  Run `register.bat` as Administrator.
2.  Logs: `%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`
