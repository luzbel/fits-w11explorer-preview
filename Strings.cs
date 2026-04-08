using System.Globalization;

namespace FitsPreviewHandler
{
    /// <summary>
    /// Centralized strongly-typed localizations.
    /// Note: For Shell Extensions (Preview Handlers), using static properties in the main assembly
    /// is safer than standard .resx satellite assemblies. 'prevhost.exe' runs from System32 and 
    /// often fails to resolve "es\FitsPreviewHandler.resources.dll" in the extension's folder.
    /// This pattern ensures zero deployment issues while keeping strings clean and separated.
    /// </summary>
    public static class Strings
    {
        private static bool IsEs => CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "es";

        public static string Title => IsEs ? "Vista Previa FITS" : "FITS Preview";
        
        public static string LogStatus(bool logOn) => IsEs ? $"Trazas: {(logOn ? "ON" : "OFF")}" : $"Trace: {(logOn ? "ON" : "OFF")}";
        
        public static string RightClickHint => IsEs 
            ? "Clic derecho para opciones de configuración y copiado" 
            : "Right-Click for configuration and copy options";
        
        public static string FileNotFound => IsEs ? "Archivo no encontrado:\n" : "File not found:\n";
        public static string ErrorOpening => IsEs ? "Error abriendo archivo FITS:\r\n" : "Error opening FITS file:\r\n";
        public static string ErrorReadingStream => IsEs ? "Error leyendo el flujo FITS:\r\n" : "Error reading FITS stream:\r\n";
        
        public static string NoImageHint => IsEs ? " · sin imagen" : " · no image";
        public static string ImageHiddenHint => IsEs ? " (Imagen oculta)" : " (Image hidden)";
        
        public static string RenderBitpix0 => IsEs ? "⚠ BITPIX=0, no se puede renderizar" : "⚠ BITPIX=0, cannot render";
        public static string RenderTooLarge(long mb) => IsEs ? $"⚠ Imagen demasiado grande ({mb} MB)" : $"⚠ Image too large ({mb} MB)";
        public static string RenderError(string err) => IsEs ? $"Error de renderizado: {err}" : $"Render Error: {err}";
        public static string RenderProgress(double pct) => IsEs ? $"Leyendo datos FITS... {pct:F0}%" : $"Reading FITS data... {pct:F0}%";
        
        // Context Menu
        public static string MenuHideImage => IsEs ? "Ocultar imagen (carga ultra-rápida). Solo tabla" : "Hide image (ultra-fast loading). Metadata only";
        public static string MenuShowImage => IsEs ? "Previsualizar imagen con auto-estirado" : "Preview image with auto-stretch";
        
        public static string MenuSavedImage => IsEs 
            ? "¡Guardado! Selecciona otro archivo FITS para ver el cambio de imagen." 
            : "Saved! Select another FITS file to apply image settings.";
            
        public static string MenuDisableTrace => IsEs ? "Desactivar trazas de depuración (Trace: OFF)" : "Disable debug tracing (Trace: OFF)";
        public static string MenuEnableTrace => IsEs ? "Activar trazas en %USERPROFILE%\\AppData\\LocalLow\\FitsPreviewHandler" : "Enable tracing in %USERPROFILE%\\AppData\\LocalLow\\FitsPreviewHandler";
        
        public static string MenuSavedTrace => IsEs 
            ? "¡Guardado! Selecciona otro archivo FITS para aplicar el nivel de log." 
            : "Saved! Select another FITS file to apply tracing settings.";
        
        public static string MenuCopyImage(int w, int h) => IsEs ? $"Copiar imagen ({w}x{h})" : $"Copy image ({w}x{h})";
        public static string MenuCopyRow => IsEs ? "Copiar fila seleccionada" : "Copy selected row";
        public static string MenuCopyCsv => IsEs ? "Copiar toda la tabla (CSV)" : "Copy entire table (CSV)";
    }
}
