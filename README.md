[English (EN)](README.md) 🇬🇧 | [Español (ES)](README.es.md) 🇪🇸

# FITS Windows Explorer Preview Handler

> **Native FITS file previewer for Windows 10/11 Explorer.**

![Platform: Windows 10/11](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)
![Zero-Copy](https://img.shields.io/badge/Architecture-Zero--Copy-green)

---

## ✨ Main Features

- **Zero-Copy Architecture**: Reads multi-GB FITS files instantly without copying to disk, using direct streams (`IStream`).
- **Thumbnail Provider**: Generates native Windows Explorer thumbnails for `.fits` files:
  - **Stride Sampling**: Only reads the fraction of pixels needed to fill the requested thumbnail size. For a 4656×3520 sensor at 256 px the stride is ≈18, so only ~1/324th of the pixel data is ever read.
  - **Non-Blocking**: The Shell calls `GetThumbnail` on a background thread per file, in parallel across a whole folder, while the Explorer UI stays responsive.
  - **Shell Cache**: Thumbnails are cached automatically in `%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache_*.db` by the OS — regeneration only happens when the file changes.
  - **Static Badge mode**: When the image panel is disabled (`ShowImage = 0`) the thumbnail never reads pixel data at all. Instead it renders a colour-coded identification badge:
    - Background and icon encode the frame type (`LIGHT ★` / `DARK ■` / `FLAT ●` / `BIAS ─`).
    - The accent colour encodes the filter (`Hα` = red, `OIII` = cyan, `SII` = orange, `Hβ` = blue, `L/R/G/B` bands = their respective colours).
    - The frame type label, filter name, **and a `FITS` format tag** are always printed as text, so the badge is fully readable without knowing any colour code.
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
    - `System.Category` (from FITS `IMAGETYP` or `FRAME` — normalised to `Light` / `Dark` / `Flat` / `Bias`)
  - **Display Modes**:
    - **InfoTip Support**: Hover your mouse over any FITS file to instantly view a native tooltip summarizing the camera, dimensions, and subject.
    - **Details Pane**: Press `Alt+Enter` or view the file Properties (Details tab) to see the metadata naturally loaded into the OS. The `register.bat` script configures the layout of these properties (`FullDetails`, `InfoTip`, `PreviewDetails`) in the registry.
  - **Advanced Search in Windows Explorer**:
    Because these files are natively registered, you can use Windows Advanced Query Syntax (AQS) in the Explorer search bar. It is recommended to use the exact `System.` keys to avoid localization issues:
    - **Find by frame type** — the most common use case:
      | Query | Result |
      | :--- | :--- |
      | `System.Category:Light` | All light frames |
      | `System.Category:Dark` | All dark frames |
      | `System.Category:Flat` | All flat fields |
      | `System.Category:Bias` | All bias frames |
    - **Find by camera or object**: `System.Photo.CameraModel:ZWO` · `System.Subject:M31`
    - **Numeric ranges**: `System.Photo.ExposureTime:>120` (over 120 s) · `System.Photo.ExposureTime:10..300`
    - **Combining conditions (AND / OR — must be UPPERCASE)**:
      ```
      System.Category:Dark AND System.Photo.ExposureTime:>300
      System.Category:Light AND System.Subject:M31
      System.Category:Flat AND System.Photo.CameraModel:"ASI2600"
      ```
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
| `ShowImage` | `1` (Show image and table + real pixel thumbnail), `0` (Table only + static badge thumbnail). |
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
