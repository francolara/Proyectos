# Fase 3 - Checklist Play Store (AppPrestamos)

Fecha: 25/03/2026

## 1) Antes de subir AAB

- Verificar build release en limpio:
  - `.\gradlew.bat :app:assembleRelease`
- Confirmar versionado:
  - `versionName` visible (ejemplo: `1.0.1`, `1.0.2`)
  - `versionCode` incremental
- Probar en dispositivo real:
  - Crear respaldo local cifrado
  - Crear respaldo en Drive cifrado
  - Restaurar con clave correcta
  - Restaurar con clave incorrecta (debe fallar con mensaje amigable)

## 2) OAuth / Drive (Google Cloud)

- Proyecto con Google Drive API habilitada.
- Pantalla OAuth configurada (nombre app, correo soporte, dominio si aplica).
- Tipo de usuarios:
  - `Externo` para cuentas Gmail normales.
- Credencial OAuth Android creada con:
  - `packageName` correcto de la app.
  - SHA-1 del keystore usado en release.
- Si OAuth está en modo prueba:
  - agregar testers (correos de prueba).
- Si OAuth está en producción:
  - publicar app OAuth en Google Auth Platform.

## 3) Play Console - Seguridad de datos (Data safety)

Declarar según comportamiento real de la app:

- Datos financieros:
  - Sí (préstamos, cuotas, pagos, montos).
- Información personal:
  - Sí (nombre, teléfono, documento).
- Datos se recopilan:
  - Sí (ingresados por usuario).
- Datos se comparten con terceros:
  - No (si no envías a backend propio ni terceros fuera de Drive del usuario).
- Cifrado en tránsito:
  - Sí (cuando usa APIs de Google/Drive).
- Opción de eliminación:
  - Sí, desde la app (eliminación de registros y restauración manual según flujo actual).

Nota:
- Si en futuro agregas analytics/crash reporting/ads, debes actualizar esta sección.

## 4) Permisos y cumplimiento

- Revisar permisos en `AndroidManifest.xml`:
  - Mantener solo los necesarios.
- No mostrar errores técnicos crudos en UI.
- Confirmar que respaldo queda cifrado por clave del usuario (portable).

## 5) Ficha de Play Store

- Título y descripción corta/larga.
- Capturas reales (modo claro y oscuro).
- Ícono 512x512.
- Banner 1024x500 (si usarás destacados).
- Categoría correcta (Finanzas/Productividad).
- Correo de contacto actualizado.

## 6) Política de privacidad (obligatoria recomendada)

Debe incluir:

- Qué datos procesa la app.
- Para qué se usan.
- Cómo funciona respaldo local y Drive.
- Que los respaldos se cifran con clave definida por usuario.
- Qué pasa si el usuario olvida su clave (no se puede descifrar respaldo).

## 7) QA final recomendado (go/no-go)

- Registro inicial y PIN.
- Licencia: prueba/activa/expirada.
- Clientes / Préstamos / Pagos / Mora.
- Reportes PDF (resumen y detallado).
- Respaldo local + Drive:
  - reemplazo de archivo existente (sin duplicados)
  - restauración correcta
  - mensajes de error amigables
- Cierre de sesión con respaldo automático:
  - con ubicación configurada
  - sin ubicación configurada

## 8) Publicación por etapas

- Subir primero a `Prueba cerrada`.
- Validar mínimo 3-5 dispositivos reales.
- Revisar crashes/ANR.
- Luego promover a producción gradual (10% -> 50% -> 100%).

