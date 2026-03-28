# SistemaControlEspaciosDeportivosWeb

## Arquitectura actual
- Frontend: ASP.NET Core MVC.
- Base de datos: SQL Server (`DbSportCenter`).
- Acceso a datos de negocio: **ADO.NET + Stored Procedures**.
- Identity (autenticacion base): ASP.NET Identity.

## Capa ADO.NET implementada
- Interfaz: `Services/ISportCenterStoredProcedureService.cs`
- Implementacion parcial:
  - `Services/SportCenterStoredProcedureService.Base.cs`
  - `Services/SportCenterStoredProcedureService.SedesEspacios.cs`
  - `Services/SportCenterStoredProcedureService.ReservasPagosComprobantes.cs`
  - `Services/SportCenterStoredProcedureService.Clientes.cs`
  - `Services/SportCenterStoredProcedureService.Reportes.cs`
  - `Services/SportCenterStoredProcedureService.Solicitudes.cs`
  - `Services/SportCenterStoredProcedureService.Usuarios.cs`
  - `Services/SportCenterStoredProcedureService.Promociones.cs`
  - `Services/SportCenterStoredProcedureService.ReservasPagosComprobantes.cs` (incluye calendario y bloqueos sprint 5)
  - `Services/SportCenterStoredProcedureService.Automatizacion.cs`
- Seguridad por modulo via SP:
  - `Services/ModuloPermisoService.cs` usa `Sp_Seguridad_ObtenerContextoModulo`.

## Controladores que ya usan ADO.NET + SP
- `Controllers/HomeController.cs`
- `Controllers/PanelController.cs`
- `Controllers/SedesController.cs`
- `Controllers/ClientesController.cs`
- `Controllers/EspaciosController.cs`
- `Controllers/ReservasController.cs`
- `Controllers/SolicitudesController.cs`
- `Controllers/UsuariosController.cs`
- `Controllers/PromocionesController.cs`
- `Controllers/PagosController.cs`
- `Controllers/ComprobantesController.cs`
- `Controllers/ReportesController.cs`

## Carpeta de Stored Procedures
`Basededatos/deStoreProcedures_DbSportCenter`

### 00_Auditoria.sql
- `Sp_Auditoria_Registrar`

### 01_Seguridad_Panel.sql
- `Sp_Seguridad_SeedModulosPermisosBase`
- `Sp_Seguridad_ObtenerContextoModulo`
- `Sp_Panel_ListarNegociosUsuario`
- `Sp_Panel_ObtenerRolUsuario`
- `Sp_Panel_ListarModulosPermitidos`
- `Sp_Panel_ObtenerMetricas`

### 02_Home.sql
- `Sp_Home_ListarSedesPublicas`
- `Sp_Home_ListarTiposDeporte`
- `Sp_Home_BuscarEspaciosDisponibles`

### 03_Sedes_Espacios.sql
- `Sp_Combos_Sedes`
- `Sp_Combos_TiposDeporte`
- `Sp_Sedes_Listar`
- `Sp_Sedes_ObtenerPorId`
- `Sp_Sedes_Crear`
- `Sp_Sedes_Actualizar`
- `Sp_Sedes_Eliminar`
- `Sp_Espacios_Listar`
- `Sp_Espacios_ObtenerPorId`
- `Sp_Espacios_Crear`
- `Sp_Espacios_Actualizar`
- `Sp_Espacios_Eliminar`

### 04_Reservas_Pagos_Comprobantes.sql
- `Sp_Combos_EspaciosPorNegocio`
- `Sp_Combos_Clientes`
- `Sp_Combos_ReservasPorNegocio`
- `Sp_Reservas_Listar`
- `Sp_Reservas_ObtenerPorId`
- `Sp_Reservas_Crear`
- `Sp_Reservas_Actualizar`
- `Sp_Reservas_Eliminar`
- `Sp_Pagos_Listar`
- `Sp_Pagos_ObtenerPorId`
- `Sp_Pagos_Crear`
- `Sp_Pagos_Actualizar`
- `Sp_Pagos_Eliminar`
- `Sp_Comprobantes_Listar`
- `Sp_Comprobantes_ObtenerPorId`
- `Sp_Comprobantes_Crear`
- `Sp_Comprobantes_Actualizar`
- `Sp_Comprobantes_Eliminar`

### 05_Sedes_Servicios.sql
- Tablas:
  - `CatalogoServiciosSede`
  - `SedeServicios`
- `Sp_Combos_ServiciosSede`
- `Sp_Sedes_Listar` (incluye columna `Servicios`)
- `Sp_Sedes_ObtenerPorId` (incluye columna `ServiciosIdsCsv`)
- `Sp_Sedes_Crear` (nuevo parametro `@ServiciosIdsCsv`)
- `Sp_Sedes_Actualizar` (nuevo parametro `@ServiciosIdsCsv`)

### 06_Seguridad_Clientes.sql
- Tabla:
  - `NegocioClientes`
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `CLIENTES`)
- `Sp_Combos_Clientes` (filtrado por `@NegocioId`)
- `Sp_Clientes_Listar`
- `Sp_Clientes_ObtenerPorId`
- `Sp_Clientes_Crear`
- `Sp_Clientes_Actualizar`
- `Sp_Clientes_Eliminar`

### 07_Reservas_Pagos_Reglas.sql
- `Sp_Reservas_Crear` (valida cruce, horas y montos)
- `Sp_Reservas_Actualizar` (valida cruce, horas y montos)
- `Sp_Pagos_Crear` (valida monto y evita sobrepago)
- `Sp_Pagos_Actualizar` (recalcula saldo en reserva nueva/anterior)
- `Sp_Pagos_Eliminar` (recalcula saldo de la reserva)

### 08_Reservas_Calendario_Filtros.sql
- `Sp_Reservas_Listar` (filtros por rango fecha, sede, espacio y estado)

### 09_Reportes_Ocupacion_Ingresos.sql
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `REPORTES`)
- `Sp_Reportes_OcupacionPorEspacio`
- `Sp_Reportes_IngresosPorDia`

### 10_Home_Solicitudes_Publicas.sql
- Tabla:
  - `SolicitudesReservaPublica`
- `Sp_Home_SolicitarReservaPublica`

### 11_Solicitudes_Gestion.sql
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `SOLICITUDES`)
- `Sp_SolicitudesPublicas_Listar`
- `Sp_SolicitudesPublicas_ActualizarEstado`
- `Sp_SolicitudesPublicas_ConvertirAReserva`

### 12_Home_Notificaciones_Seguimiento.sql
- `Sp_Home_ConsultarSolicitudPublica`
- `Sp_Home_ObtenerSolicitudParaNotificacion`
- `Sp_Home_MarcarSolicitudNotificada`

### 13_Usuarios_Negocio_Gestion.sql
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `USUARIOS`)
- `Sp_UsuariosNegocio_Listar`
- `Sp_UsuariosNegocio_AsignarPorCorreo`
- `Sp_UsuariosNegocio_ActualizarRol`
- `Sp_UsuariosNegocio_Desactivar`
- `Sp_UsuariosNegocio_PermisosListar`
- `Sp_UsuariosNegocio_PermisoGuardar`

### 14_Promociones_Kpis.sql
- Tabla:
  - `PromocionesHorario`
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `PROMOCIONES`)
- `Sp_Panel_ObtenerMetricas` (agrega `OcupacionHoyPct`, `NoShowMes`, `TicketPromedioMes`)
- `Sp_Promociones_Listar`
- `Sp_Promociones_ObtenerPorId`
- `Sp_Promociones_Crear`
- `Sp_Promociones_Actualizar`
- `Sp_Promociones_Eliminar`

### 15_Calendario_Bloqueos.sql
- Tabla:
  - `BloqueosHorario`
- `Sp_Reservas_CalendarioEventos`
- `Sp_Reservas_Mover`
- `Sp_Bloqueos_Listar`
- `Sp_Bloqueos_Crear`
- `Sp_Bloqueos_Eliminar`

### 16_Reservas_CheckIn_CheckOut.sql
- `Sp_Reservas_CambiarEstadoRapido`
- Reglas:
  - `Check-in` (3) desde `Pendiente` (1) o `Confirmada` (2)
  - `Check-out` (4) solo desde `En uso` (3)
  - `No-show` (6) desde `Pendiente` (1) o `Confirmada` (2)

### 17_Automatizacion_Recordatorios_NoShow.sql
- Alter tabla `Reservas`:
  - `RecordatorioEnviado` (BIT)
  - `FechaRecordatorio` (DATETIME2)
- `Sp_Reservas_RecordatoriosPendientes`
- `Sp_Reservas_MarcarRecordatorioEnviado`
- `Sp_Reservas_AutoNoShow`

### 18_Sedes_Config_Notificaciones.sql
- Tabla:
  - `SedeConfiguracionNotificacion`
- Ajusta SP de sedes para guardar configuracion por sede:
  - `Sp_Sedes_Listar`
  - `Sp_Sedes_ObtenerPorId`
  - `Sp_Sedes_Crear`
  - `Sp_Sedes_Actualizar`
- Ajusta automatizacion por sede:
  - `Sp_Reservas_RecordatoriosPendientes` (usa anticipacion por sede)
  - `Sp_Reservas_AutoNoShow` (usa tolerancia por sede)

### 19_Home_Whatsapp_Publico.sql
- Ajusta `Sp_Home_ListarSedesPublicas` para devolver:
  - `WhatsappContacto`
  - `PermiteChatWhatsapp`
- Permite mostrar boton "Chatear por WhatsApp" en el portal publico por sede.

### 20_Home_Espacios_Whatsapp.sql
- Ajusta `Sp_Home_BuscarEspaciosDisponibles` para devolver:
  - `WhatsappContacto`
  - `PermiteChatWhatsapp`
- Permite mostrar boton "Chatear por WhatsApp" directamente en tarjetas de espacios disponibles.

### 21_Altas_Clubes.sql
- Tabla:
  - `SolicitudesAltaClub`
- Flujo publico:
  - `Sp_Home_SolicitarAltaClub`
- Gestion interna:
  - `Sp_AltasClubes_Listar`
  - `Sp_AltasClubes_Aprobar`
  - `Sp_AltasClubes_Rechazar`
- Al aprobar:
  - crea `Negocio`
  - crea primera `Sede`
  - intenta vincular usuario existente por correo como `RolNegocio = 1`

### 22_Registro_Club_Prueba.sql
- Tabla:
  - `NegociosSuscripcion`
- Flujo publico directo:
  - `Sp_Home_RegistrarClubConPrueba`
- Al registrar:
  - crea `Negocio`
  - crea primera `Sede`
  - asocia al usuario nuevo como `RolNegocio = 1`
  - activa prueba automatica de 30 dias
  - registra alta como solicitud autoaprobada (si existe `SolicitudesAltaClub`)

### 23_Suscripcion_Bloqueo_Operacion.sql
- Reglas de bloqueo:
  - auto cambia a `Vencida` cuando la prueba o plan supera su fecha fin.
  - bloquea acceso a modulos cuando estado de suscripcion es `Vencida` o `Suspendida`.
- SP actualizados:
  - `Sp_Seguridad_ObtenerContextoModulo`
  - `Sp_Panel_ListarModulosPermitidos`
- SP nuevo:
  - `Sp_NegociosSuscripcion_ActivarPlan` (reactiva acceso y define vigencia en dias)

## Flujo recomendado de despliegue SQL
1. Ejecutar `00_Auditoria.sql`.
2. Ejecutar `01_Seguridad_Panel.sql`.
3. Ejecutar `02_Home.sql`.
4. Ejecutar `03_Sedes_Espacios.sql`.
5. Ejecutar `04_Reservas_Pagos_Comprobantes.sql`.
6. Ejecutar `05_Sedes_Servicios.sql`.
7. Ejecutar `06_Seguridad_Clientes.sql`.
8. Ejecutar `07_Reservas_Pagos_Reglas.sql`.
9. Ejecutar `08_Reservas_Calendario_Filtros.sql`.
10. Ejecutar `09_Reportes_Ocupacion_Ingresos.sql`.
11. Ejecutar `10_Home_Solicitudes_Publicas.sql`.
12. Ejecutar `11_Solicitudes_Gestion.sql`.
13. Ejecutar `12_Home_Notificaciones_Seguimiento.sql`.
14. Ejecutar `13_Usuarios_Negocio_Gestion.sql`.
15. Ejecutar `14_Promociones_Kpis.sql`.
16. Ejecutar `15_Calendario_Bloqueos.sql`.
17. Ejecutar `16_Reservas_CheckIn_CheckOut.sql`.
18. Ejecutar `17_Automatizacion_Recordatorios_NoShow.sql`.
19. Ejecutar `18_Sedes_Config_Notificaciones.sql`.
20. Ejecutar `19_Home_Whatsapp_Publico.sql`.
21. Ejecutar `20_Home_Espacios_Whatsapp.sql`.
22. Ejecutar `21_Altas_Clubes.sql`.
23. Ejecutar `22_Registro_Club_Prueba.sql`.
24. Ejecutar `23_Suscripcion_Bloqueo_Operacion.sql`.
25. Ejecutar `EXEC dbo.Sp_Seguridad_SeedModulosPermisosBase;` una vez.

## Observaciones funcionales
- CRUD de modulos internos ejecuta operaciones por SP.
- Validacion de permisos `PuedeVer/PuedeCrear/PuedeEditar/PuedeEliminar` se resuelve por SP de seguridad.
- Auditoria de acciones CRUD se registra con `Sp_Auditoria_Registrar`.
- Home publica permite registro de solicitudes de reserva con codigo de seguimiento.
- Notificacion por correo configurable via `EmailSettings` (SMTP).
- Modulo de promociones permite descuentos por sede/espacio con vigencia por fecha y franja horaria.
- Panel privado muestra KPIs avanzados para operacion diaria y seguimiento mensual.
- Reservas integra FullCalendar con vista semana/dia/mes, arrastre para mover horarios y bloqueos operativos por espacio.
- Reservas permite cambio rapido de estado (check-in/check-out/no-show) desde tabla y calendario.
- Automatizacion en segundo plano:
  - envia recordatorios por correo antes de la hora de reserva
  - marca no-show automatico segun tolerancia configurada
- Configuracion por sede en formulario de Sedes:
  - activar/desactivar notificaciones
  - minutos de anticipacion del recordatorio
  - minutos de tolerancia para no-show automatico
  - correo de notificacion del negocio (copia oculta en recordatorios)
  - numero de WhatsApp de contacto para chatear/coordinar
- Portal publico:
  - boton de WhatsApp visible solo si la sede habilito chat y registro numero.
  - boton tambien visible en tarjetas de espacios disponibles para iniciar chat inmediato.
- Boton publico "Software para Clubes":
  - pide contrasena para crear la cuenta del dueno
  - valida codigo CAPTCHA
  - crea negocio/sede en el registro inicial
  - activa prueba automatica de 1 mes
- Control de suscripcion:
  - al vencer prueba o plan, el negocio queda bloqueado para operar modulos.
  - para reactivar se usa `Sp_NegociosSuscripcion_ActivarPlan`.
