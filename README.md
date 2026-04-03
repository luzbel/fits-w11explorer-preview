# fits-w11explorer-preview

> **Windows 11 Explorer preview handler for FITS astronomical image files**

![License: CC BY-NC 4.0](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey.svg)
![Platform: Windows 11](https://img.shields.io/badge/Platform-Windows%2011-blue)
![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-purple)

---

## What it does

When you select a `.fits` file in **Windows Explorer**, the preview pane on the right shows a formatted table with the complete **FITS primary header**: keywords, values and comments — without opening any external application.

| Keyword colour | Meaning |
|---|---|
| 🟡 Yellow | `NAXIS*` — image dimensions |
| 🟢 Green | `COMMENT` / `HISTORY` records |
| 🔴 Red | `END` — header terminator |

## Requirements

| Requirement | Version |
|---|---|
| Windows | 10 / 11 (64-bit) |
| .NET Framework | 4.8 |
| Visual Studio / dotnet CLI | Any recent version |

## Build

```powershell
dotnet build FitsPreviewHandler.csproj -c Debug
```

Output: `bin\Debug\net48\FitsPreviewHandler.dll`

## Installation

> ⚠️ Run PowerShell **as Administrator**.

```powershell
.\Register-Extension-Fix.ps1
```

This script registers the COM component and associates it with the `.fits` extension in the Windows Registry.

After registration, **restart Explorer** or log off/on to activate the preview handler.

## How the FITS parser works

The handler implements the standard Windows Preview Handler COM interfaces:

- `IPreviewHandler` — lifecycle (DoPreview / Unload / SetWindow / SetRect)
- `IInitializeWithFile` — receives the file path from Explorer
- `IInitializeWithStream` — fallback stream-based init (copies to a temp file)
- `IInitializeWithItem` — fallback shell item init

The FITS header is parsed following the official FITS standard:
- Header blocks are **2880 bytes** (36 records × 80 bytes each)
- String values handle embedded `''` (escaped single quote)
- Comments after `/` are correctly separated from values

## Debugging

The handler writes a detailed trace log next to the DLL:

```
bin\Debug\net48\fits_trace.log
```

Follow it live from PowerShell (no tools needed):

```powershell
Get-Content "bin\Debug\net48\fits_trace.log" -Wait
```

Each line includes timestamp, context (`[Extension]` or `[Control]`) and thread ID, so you can trace exactly what Explorer called and when.

## Project structure

```
FitsPreviewHandler.csproj          — .NET 4.8 project
FitsPreviewHandlerExtension.cs     — COM interfaces + Preview Handler class
FitsPreviewControl.cs              — WinForms UserControl (UI + FITS parser)
TestApp.cs                         — Standalone test host (no COM needed)
Register-Extension-Fix.ps1         — Registration script (run as Admin)
Register-Extension.ps1             — Alternative registration script
```

## License

[CC BY-NC 4.0](LICENSE) — Free to use, study and modify. **Commercial use is not permitted.**
