# Etapa 2 — acceso y presentación HYT

03/11/2026. Implementación y verificación local completadas, con integración SQL pendiente de ejecución manual por el usuario.

## Entregado

- Acceso privado /Acceso, confirmación de salida /Salir y acceso restringido /Denegado.
- Hashes Identity, cookies seguras, sello de seguridad, permisos, revalidación de sesión por solicitud, bloqueo por intentos y protección antifalsificación.
- Repositorio ADO.NET y cinco SP: consultar identidad/roles/permisos, finalizar intento y auditar, registrar salida, crear administrador inicial y restablecer acceso. Script de permisos web limitado.
- Herramienta manual Hyt.Admin para administrador inicial y recuperación asistida, sin cuentas de fábrica ni secretos en archivos.
- Menú con permisos, agrupación y nombres HYT, adaptable a celular; rutas de negocio honestamente pendientes.
- Controladores, ViewModels y vistas MVC; parciales de grilla, filtros, ficha, popup, calendario y trazabilidad. Identidad HYT y calendario de FralseCont adaptado a JavaScript MVC.
- Paquete incremental ETAPA_2_HYT.sql con seis scripts y SHA256_ETAPA_2.json; aplicar después de etapa 1.

## Evidencia

Compilación sin errores/advertencias; 29 comprobaciones HTTP de autenticación, autorización, exclusividad MVC y parciales sobre el host real con repositorio de pruebas. Chrome revisado en escritorio/móvil, sin desbordamiento; módulo nativo de popup con Escape/retorno de foco y bloqueo ocupado. Capturas de acceso conservadas. Todos los scripts analizados offline; no se conectó SQL Server ni se crearon usuarios reales.

La configuración y pasos de activación están en GUIA_ACCESO_ETAPA_2.md. No se realizó despliegue, prueba de credenciales reales, email automático, CRUD de usuarios/roles o módulos de negocio. El siguiente trabajo corresponde a los mantenimientos de etapa 3.

Ajuste solicitado el 03/11/2026: retirada la interfaz Blazor de la solución activa. Los scripts y la tecnología de acceso a datos se mantienen: C#, ADO.NET y procedimientos almacenados.
