[English (EN)](README.md) 🇬🇧 | [Español (ES)](README.es.md) 🇪🇸

# FITS Windows Explorer Preview Handler

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
- **Windows Property System Integration**: Extracts metadata from FITS headers and injects them natively into Windows Explorer. *Note: Only specific properties are mapped because they must correspond to standard Windows System Properties inside the C# code. Custom FITS keywords cannot be arbitrarily indexed without modifying the source code.*
  - **Mapped & Indexed Properties**:
    - `System.Subject` (mapped from FITS `OBJECT`)
    - `System.Image.HorizontalSize` (from `NAXIS1`)
    - `System.Image.VerticalSize` (from `NAXIS2`)
    - `System.Image.BitDepth` (from `BITPIX`)
    - `System.Photo.CameraModel` (from FITS `INSTRUME` or `CAMERA`)
    - `System.Photo.ExposureTime` (from FITS `EXPOSURE` or `EXPTIME`)
  - **Display Modes**:
    - **InfoTip Support**: Hover your mouse over any FITS file to instantly view a native tooltip summarizing the camera, dimensions, and subject.
    - **Details Pane**: Press `Alt+Enter` or view the file Properties (Details tab) to see the metadata naturally loaded into the OS. The `register.bat` script configures the layout of these properties (`FullDetails`, `InfoTip`, `PreviewDetails`) in the registry.
  - **Advanced Search in Windows Explorer**:
    Because these files are natively registered, you can use Windows Advanced Query Syntax (AQS) in the Explorer search bar. It is recommended to use the exact `System.` keys to avoid localization issues:
    - **Exact text match**: `System.Photo.CameraModel:ZWO` or `System.Subject:M31`
    - **Numeric Ranges (Greater/Less than)**: Filter by exposure boundaries, for example: `System.Photo.ExposureTime:>120` (more than 120s) or `System.Photo.ExposureTime:10..300` (between 10s and 300s).
    - **Combining Conditions (AND / OR)**: Chain queries using logical operators (must be UPPERCASE).
      *Example: `System.Photo.CameraModel:"ASI294" AND System.Photo.ExposureTime:>300`*
- **Dynamic Layout**: Image on top and metadata table at the bottom with resizable height.
- **Progressive Streaming**: Displays reading progress over the image area while processing the data stream.
- **Context Menu Integration**: Right-click anywhere in the preview window to access settings and export options:
  - **Toggle Image / Metadata Only**: Instantly configure whether to render the image or load only the metadata grid for maximum speed.
  - **Toggle Tracing Logs**: Enable or disable debug mode on the fly.
  - **Copy Image**: Copies the auto-stretched and scaled preview image directly to your clipboard (ready for Paint/Photoshop).
  - **Copy Selected Row**: Copies the selected FITS keyword row to the clipboard.
  - **Copy Full Table (CSV)**: Export the entire FITS header into a comma-separated format with headers.

## ⚙️ Configuration (Registry)

Configuration is now seamlessly managed via the **Right-Click Context Menu**, and changes are applied instantly upon selecting the next file. No external executable is required.

However, if you want to automate deployment, you can edit the Windows Registry. Settings are securely stored in the Low-Integrity zone:

### User Level (Recommended)
`HKEY_CURRENT_USER\Software\AppDataLow\FitsPreviewHandler`

### Global Level (Default)
`HKEY_LOCAL_MACHINE\Software\AppDataLow\FitsPreviewHandler`

| Value (DWORD) | Description |
| :--- | :--- |
| `ShowImage` | `1` (Show image and table), `0` (Table only). |
| `EnableTracing` | `1` (Enable debug logs), `0` (Disable). |

---

## 🚀 Installation (As Administrator)

1.  Build the project using `dotnet build -c Release`.
2.  Run `register.bat` to register the component in the system.
3.  If you need to uninstall, use `unregister.bat`.

---

## 🔍 Diagnostics (Logs)

Logs are saved in the Low-Integrity area to comply with `prevhost.exe` sandbox restrictions:
`%USERPROFILE%\AppData\LocalLow\FitsPreviewHandler\fits_trace.log`
