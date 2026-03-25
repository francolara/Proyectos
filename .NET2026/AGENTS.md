# Instrucciones Codex para el repositorio Proyectos

## Alcance general
1. Revisa estas pautas antes de trabajar en cualquier archivo de este repositorio.
2. Limita los cambios a la lógica de negocio. La interfaz gráfica de los proyectos WinForms (VB6 o .NET) no debe alterarse.

## Codificaciones obligatorias
- VisualBasic 6/** (archivos .bas, .cls, .frm, .vbp, .vbw):
  - Abrir/guardar .frm siempre como Windows-1252 (ANSI) **sin BOM** y finales de línea CRLF.
  - Evitar editores que auto-conviertan a UTF-8 sin control.
  - Revisa que los acentos y caracteres especiales conserven su grafía original; evita que el editor convierta a UTF-8 o normalice saltos de línea.
  - Si el archivo queda en UTF-8 tras aplicar un parche, reconvierte con:
    iconv -f UTF-8 -t WINDOWS-1252 archivo > tmp && mv tmp archivo
  - No modifiques los binarios .frx.
- Scripts/** (archivos .sql, .ddl, .dml, etc.):
  - Para todos los `.sql` usa **UTF-8 sin BOM** y finales de línea **LF**. Verifica con `file` (o herramienta equivalente) antes de confirmar cambios.
  - Verifica en tu editor o con `file` antes de confirmar cambios.

## Checklist previo a confirmar
1. Ejecuta `git status` y `git diff --stat` para comprobar que solo cambiaste lo necesario.
2. Usa `file "ruta/al/archivo"` para validar la codificación de cada archivo modificado.
3. Asegúrate de que los formularios VB6 sigan mostrando únicamente modificaciones de código backend.

## Otros recordatorios
- Los procedimientos, scripts y tablas residen en la carpeta Basededatos/.
- Para cambios masivos de codificación, apóyate en `.gitattributes` y en `git add --renormalize .`.