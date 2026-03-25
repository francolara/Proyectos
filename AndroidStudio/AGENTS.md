## Codificación obligatoria (Android/Kotlin)

1. Todos los archivos .kt, .kts, .xml, .properties, .md deben guardarse en **UTF-8 sin BOM**.
2. No convertir archivos a ANSI/Windows-1252.
3. No cambiar finales de línea masivamente: mantener los existentes del repo.
4. Si se detecta mojibake (Ã, Â, ðŸ, â€, âœ), corregir antes de finalizar.
5. Para emojis en Kotlin, preferir escapes Unicode (\uD83D...) en strings sensibles.
6. Después de editar, validar:
   - búsqueda de patrones corruptos en app/src/main
   - compilación :app:compileDebugKotlin
7. Antes de cerrar la tarea, confirmar que los archivos editados estén en UTF-8 sin BOM.

## Regla de edición segura
- Evitar reemplazos masivos con herramientas que cambien codificación.
- En PowerShell, no usar Set-Content -Encoding UTF8 (puede meter BOM según versión).
- Usar escritura explícita UTF-8 sin BOM cuando se edite por script.