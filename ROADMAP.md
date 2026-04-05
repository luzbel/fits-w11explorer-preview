# FITS Preview Handler - Roadmap/Pendientes

Este documento lista las mejoras planeadas para el previsualizador de archivos FITS.

## 🚀 Próximos Pasos (Core)
- [x] **Lectura Zero-Copy**: Eliminar la copia a archivos temporales para permitir previsualización instantánea de archivos de varios GB.
- [ ] **Gestión de Logs**: Implementar un sistema de purga automática o rotación para que los archivos en `LocalLow` no crezcan indefinidamente. Valorar el uso de `EventLog` para errores críticos.

## 📊 Vista de Detalles (Property Handler)
El objetivo es que Windows Explorer muestre metadatos del FITS en las columnas de la vista "Detalles" y permita la búsqueda por estos campos.
- **Campos a implementar**:
  - `INSTRUME`: Cámara utilizada.
  - `FILTER`: Filtro astronómico.
  - `OBJNAME`: Objeto fotografiado.
  - `FOCALLEN`: Distancia focal.
  - `EXPTIME`: Tiempo de exposición.
- **Tecnología**: Requiere la implementación de la interfaz `IPropertyStore`.

## 🎨 Mejoras de UI
- **Auto-Stretch**: Implementar algoritmos de estirado automático (Histogram Transformation) similares a Siril/PI para objetos de cielo profundo que vienen muy oscuros.
- **Soporte de Color**: Mejorar la detección de patrones Bayer desde las cabeceras `BAYERPAT`.

## 🛠️ Mantenimiento
- **CFITSIO**: Valorar la migración a la librería estándar `cfitsio` si las cabeceras FITS se vuelven demasiado complejas para el parser actual.
