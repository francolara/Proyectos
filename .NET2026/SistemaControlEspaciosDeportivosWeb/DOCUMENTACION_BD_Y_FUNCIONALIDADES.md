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
- `Services/SportCenterStoredProcedureService.Maestros.cs`
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
- `Controllers/MaestrosController.cs`

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
- Actualizacion 13/04/2026:
  - `Sp_Espacios_Listar` compacta `TarifaResumen` por dia de semana con rango de precios (`min-max`) y elimina el detalle por cada franja horaria en el listado de espacios.
  - `Sp_Espacios_Listar` expone `TieneIluminacion` y `Techada` para mostrar badges operativos en la grilla.
- `Sp_Sedes_Eliminar` y `Sp_Espacios_Eliminar` ahora retornan error cuando no existe el registro para el negocio.

### 04_Reservas_Pagos_Comprobantes.sql
- `Sp_Combos_EspaciosPorNegocio`
- `Sp_Combos_Clientes`
- `Sp_Combos_ReservasPorNegocio`
- `Sp_Reservas_Listar`
- `Sp_Reservas_ObtenerPorId`
- `Sp_Reservas_Crear`
- Actualizacion 14/04/2026:
  - `Reservas` incorpora columna `CanalOrigen` (`ADMIN`/`CLIENTE_WEB`), usada para distinguir reservas creadas desde portal cliente.
  - `Sp_Reservas_Crear` recibe `@CanalOrigen` (default `ADMIN`) y genera notificacion cuando el origen es `CLIENTE_WEB`.
- `Sp_Reservas_Actualizar`
- `Sp_Reservas_Eliminar`
- `Sp_Reservas_Eliminar` valida `@NegocioId` por join con `Sedes` y devuelve error si no encuentra la reserva.
- `Sp_Reservas_Eliminar` bloquea cancelar cuando la reserva tiene pagos registrados (`RAISERROR` para validacion funcional).
- `Sp_Pagos_Listar`
- `Sp_Pagos_ObtenerPorId`
- `Sp_Pagos_Crear`
- `Sp_Pagos_Actualizar`
- `Sp_Pagos_Eliminar`
- `Sp_Pagos_Actualizar` y `Sp_Pagos_Eliminar` devuelven error si no existe el pago para el negocio (evita falso positivo en C#).
- Actualizacion 13/04/2026:
  - `Sp_Pagos_Listar` agrega rango opcional `@FechaDesde/@FechaHasta` (fecha de reserva) para filtros rapidos en UI.
- `Sp_Comprobantes_Listar`
- `Sp_Comprobantes_Listar` incluye filtro opcional por `CodigoDocumento` (tipo de documento SUNAT del comprobante) para listar por negocio segun tipos configurados en maestros.
- Actualizacion 13/04/2026:
  - `Sp_Comprobantes_Listar` agrega rango opcional `@FechaDesde/@FechaHasta` (fecha de emision) para filtros rapidos en UI.
- `Sp_Comprobantes_ObtenerPorId`
- `Sp_Comprobantes_Crear`
- `Sp_Comprobantes_Actualizar`
- `Sp_Comprobantes_Eliminar`
- `Sp_Comprobantes_Actualizar` y `Sp_Comprobantes_Eliminar` devuelven error si no existe el comprobante para el negocio.
- `Sp_Comprobantes_Eliminar` marca el comprobante como anulado (`Estado = 5`) y libera la reserva asociada para permitir nueva emision de comprobante sobre reservas pagadas.
- `Sp_ParametrosGlobales_ObtenerValor` retorna `ValorParametro` por `NombreParametro` para reglas de validacion configurables.

### Parametros globales
- Tabla:
  - `ParametrosGlobales` (`ParametroId`, `NombreParametro`, `Descripcion`, `ValorParametro`)
- Parametro inicial:
  - `NombreParametro = VALIDA_MONTO_BSINDOC`
  - `Descripcion = Monto Maximo para atencion de boletas sin DOC`
  - `ValorParametro = 700`
- Uso funcional:
  - Validacion de comprobantes (`Boleta`) obtiene el `ValorParametro` por `NombreParametro` (sin valor fijo en codigo).

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
- `Sp_Clientes_Actualizar` y `Sp_Clientes_Eliminar` devuelven error si no existe el cliente para el negocio.
- `Sp_Clientes_Crear` y `Sp_Clientes_Actualizar` validan duplicado por `NumeroDocumento` dentro del mismo `NegocioId` (si el documento fue informado) y retornan: `Cliente ya se encuentra registrado.`

### 07_Reservas_Pagos_Reglas.sql
- `Sp_Reservas_Crear` (valida cruce, horas y montos)
- `Sp_Reservas_Actualizar` (valida cruce, horas y montos; retorna error si no afecta filas)
- `Sp_Reservas_Actualizar` bloquea cambio a `Cancelada` cuando la reserva tiene pagos registrados.
- `Sp_Pagos_Crear` (valida monto y evita sobrepago)
- `Sp_Pagos_Actualizar` (recalcula saldo en reserva nueva/anterior)
- `Sp_Pagos_Eliminar` (recalcula saldo de la reserva)
- `Sp_Pagos_Actualizar` y `Sp_Pagos_Eliminar` retornan error cuando el pago no existe para el negocio.

### 08_Reservas_Calendario_Filtros.sql
- `Sp_Reservas_Listar` (filtros por rango fecha, sede, espacio y estado simple `@Estado` o multiple `@EstadosCsv`)
- Actualizacion 13/04/2026:
  - `Sp_Reservas_Listar` calcula `SaldoPendiente` con pagos acumulados por reserva (`Total - SUM(Pagos.Monto)`, minimo 0), por lo que refleja adelantos y pagos posteriores.

### 09_Reportes_Ocupacion_Ingresos.sql
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `REPORTES`)
- `Sp_Reportes_OcupacionPorEspacio`
- `Sp_Reportes_IngresosPorDia`
- `Sp_Reportes_ResumenOperativo`
- Actualizacion 13/04/2026:
  - `Sp_Reportes_OcupacionPorEspacio` expone `SedeId` y `EspacioDeportivoId` para drill-down desde UI de reportes.
  - `Sp_Reportes_ResumenOperativo` resume estados de reserva, monto reservado/cobrado y saldo pendiente por rango/sede.
  - KPI de reportes excluye reservas canceladas (`Estado = 5`) en `TotalReservas`, `MontoReservado`, `MontoCobrado`, `SaldoPendiente` y en `Sp_Reportes_IngresosPorDia` (conteo/ingresos por dia).
  - UI de reportes agrega `Exportar Excel (.xlsx)` con hojas separadas `Resumen`, `Ocupacion` e `Ingresos` y columnas de analitica (ticket/cobranza).

### 10_Home_Solicitudes_Publicas.sql
- Tabla:
  - `SolicitudesReservaPublica`
- `Sp_Home_SolicitarReservaPublica`

### 11_Solicitudes_Gestion.sql
- `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `SOLICITUDES`)
- `Sp_SolicitudesPublicas_Listar`
- `Sp_SolicitudesPublicas_ActualizarEstado`
- `Sp_SolicitudesPublicas_ConvertirAReserva`
- `Sp_SolicitudesPublicas_ActualizarEstado` devuelve error si la solicitud no existe o no pertenece al negocio.

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
- `Sp_UsuariosNegocio_Desactivar` retorna error si no existe el usuario del negocio.
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
- `Sp_Promociones_Actualizar` y `Sp_Promociones_Eliminar` devuelven error si no existe la promoción para el negocio.
- `Sp_Promociones_Listar` (13/04/2026) incorpora filtros por `FechaDesde/FechaHasta`, estado (`@SoloActivos`: activos/inactivos/todos) y paginación (`@Pagina`, `@TamanoPagina`, `@TotalRegistros OUTPUT`).

### 15_Calendario_Bloqueos.sql
- Tabla:
  - `BloqueosHorario`
- `Sp_Reservas_CalendarioEventos`
- Actualizacion 13/04/2026:
  - `Sp_Reservas_CalendarioEventos` excluye por defecto reservas `Cancelada` (`Estado = 5`) cuando `@Estado` es `NULL`, liberando el horario en calendario para nuevas reservas.
  - Si la UI filtra explicitamente `@Estado = 5`, el procedimiento mantiene la consulta de canceladas.
- `Sp_Reservas_Mover`
- `Sp_Bloqueos_Listar`
- `Sp_Bloqueos_Crear`
- `Sp_Bloqueos_Eliminar`

### 37_Notificaciones_ReservasWeb.sql
- Tabla:
  - `NegocioNotificaciones`
- `Sp_Notificaciones_Crear`
- `Sp_Notificaciones_Listar`
- `Sp_Notificaciones_ContarNoLeidas`
- Actualizacion 14/04/2026:
  - Campanita en barra admin consulta cada 20 segundos (ajustable a 30) y muestra acumulado de notificaciones no leidas por negocio.
  - `Sp_SolicitudesPublicas_ConvertirAReserva` crea notificacion de tipo `RESERVA_CLIENTE_WEB` al generar reserva desde solicitud de portal cliente.

### 16_Reservas_CheckIn_CheckOut.sql
- `Sp_Reservas_CambiarEstadoRapido`
- `Sp_Reservas_CambiarEstadoRapido` retorna error si la reserva no existe para el negocio.
- `Sp_Reservas_CambiarEstadoRapido` bloquea cambio rapido a `Cancelada` si la reserva tiene pagos registrados.
- Reglas:
  - `Confirmada` (2) por cambio rapido
  - `Pagada` (4) desde `Pendiente` (1), `Confirmada` (2) o `En uso` historico (3)
  - `Cancelada` (5) por cambio rapido
  - `No Asistio` (6) desde `Pendiente` (1), `Confirmada` (2) o `En uso` historico (3)
  - El estado `En uso`/`Check-in` queda retirado para nuevas transiciones.

### 17_Automatizacion_Recordatorios_NoShow.sql
- Alter tabla `Reservas`:
  - `RecordatorioEnviado` (BIT)
  - `FechaRecordatorio` (DATETIME2)
- `Sp_Reservas_RecordatoriosPendientes`
- `Sp_Reservas_MarcarRecordatorioEnviado`
- `Sp_Reservas_AutoNoShow`
- `Sp_Reservas_MarcarRecordatorioEnviado` devuelve error si la reserva no existe para el negocio.

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
- Actualizacion 14/04/2026:
  - `Sp_Home_BuscarEspaciosDisponibles` cambia filtros publicos: ya no usa `SedeId`; filtra por `CodigoDepartamento`, `CodigoProvincia`, `CodigoUbigeo` y `TipoDeporteId`.
  - El resultado agrega direccion de sede, departamento/provincia/distrito, tipo de suelo y `TarifaDesde` (minimo de tarifas activas por espacio) para tarjetas del portal cliente.

### 21_Altas_Clubes.sql
- Tabla:
  - `SolicitudesAltaClub`
- Flujo publico:
  - `Sp_Home_SolicitarAltaClub`
- Gestion interna:
  - `Sp_AltasClubes_Listar`
  - `Sp_AltasClubes_Aprobar`
  - `Sp_AltasClubes_Rechazar`
- `Sp_AltasClubes_Rechazar` devuelve error cuando la solicitud no existe o ya fue gestionada.
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

### 24_Sedes_Horario_NoLaborable.sql
- Tablas:
  - `SedeHorarioAtencion`
  - `SedeFechasInhabilitadas`
- SP actualizados:
  - `Sp_Sedes_Listar`
  - `Sp_Sedes_ObtenerPorId`

### 25_Sedes_Horario_Crear_Actualizar.sql
- SP actualizados:
  - `Sp_Sedes_Crear`
  - `Sp_Sedes_Actualizar`
- Incluye persistencia de horario, servicios, notificaciones y fechas no laborables.

### 26_Reservas_Validacion_Horario_Sede.sql
- SP actualizados:
- `Sp_Reservas_Crear`
- `Sp_Reservas_Mover`
- Valida horario/dias de atencion y fechas no laborables por sede.
- `Sp_Reservas_Mover` retorna error si no encuentra la reserva o no afecta filas.

### 27_Calendario_No_Atencion_Sede.sql
- SP actualizado:
  - `Sp_Reservas_CalendarioEventos`
- Incluye bloques `NO_ATENCION` por dia no laborable y fuera de horario.
- Devuelve metadatos backend para UI:
  - `Motivo`
  - `EstadoCodigo`
  - `EstadoTexto`
- El calendario queda backend-driven para estado y motivo (sin fallback de reglas en frontend).

### 28_Espacios_Deporte_Suelo_Catalogos.sql
- Catalogos:
  - `TiposDeporte` (seed base)
  - `TiposSuelo` (seed base)
- SP actualizados:
  - `Sp_Combos_TiposDeporte`
  - `Sp_Combos_TiposSuelo`
  - `Sp_Espacios_Listar`
  - `Sp_Espacios_ObtenerPorId`
  - `Sp_Espacios_Crear`
  - `Sp_Espacios_Actualizar`

### 29_Reservas_ValidarDisponibilidad_Modal.sql
- SP:
  - `Sp_Reservas_ValidarDisponibilidad`
- Usado por modal de reservas para validar colision y restricciones de sede.

### 30_Configuracion_Club_Monedas.sql
- Catalogo:
  - `MonedasSistema`
- SP:
  - `Sp_Combos_Monedas`
  - `Sp_ConfiguracionClub_Obtener`
  - `Sp_ConfiguracionClub_Actualizar`
- Actualizacion 13/04/2026:
  - `Negocios` agrega columna `LogoUrl` para persistir el logo del club/negocio.
  - `Sp_ConfiguracionClub_Obtener` expone `LogoUrl`.
  - `Sp_ConfiguracionClub_Actualizar` permite guardar/remover `LogoUrl`.
  - script incremental: `Basededatos/SportCenter/Script/20260413_Negocios_LogoUrl.sql`.

### 31_Espacios_Tarifas_Base.sql
- Tabla:
  - `Tarifas`
- SP actualizados:
  - `Sp_Espacios_ObtenerPorId`
  - `Sp_Espacios_Crear`
  - `Sp_Espacios_Actualizar`

### 32_Usuarios_Sede_Restriccion_Filtros.sql
- Alter tabla:
  - `UsuariosNegocio.SedeId` (FK a `Sedes`)
- SP actualizados:
  - `Sp_Seguridad_ObtenerContextoModulo` (devuelve `SedeIdAsignada` y `EsAdministrador`)
  - `Sp_UsuariosNegocio_Listar` (nuevo parametro `@SedeId` para filtrar en backend)
  - `Sp_UsuariosNegocio_AsignarPorCorreo`
  - `Sp_UsuariosNegocio_ActualizarRol`
  - `Sp_Combos_Sedes`
  - `Sp_Sedes_Listar`
  - `Sp_Espacios_Listar`
  - `Sp_Combos_EspaciosPorNegocio`
  - `Sp_Combos_ReservasPorNegocio`
  - `Sp_Pagos_Listar`
  - `Sp_Comprobantes_Listar`
  - `Sp_Reportes_OcupacionPorEspacio`
  - `Sp_Reportes_IngresosPorDia`
  - `Sp_Reportes_ResumenOperativo`
  - `Sp_Panel_ObtenerMetricas`
  - `Sp_Promociones_Listar`
- Regla funcional:
  - usuarios no administradores trabajan con una sola sede asignada y los filtros/listados se restringen en backend.
  - `Sp_Panel_ObtenerMetricas` (13/04/2026) calcula `OcupacionHoyPct` con **horas disponibles netas** por dia: horario de atencion de sede menos bloqueos activos (`BloqueosHorario`) y excluye fechas inhabilitadas de sede.
  - `Sp_Espacios_Listar` devuelve `TarifaResumen` (dias + rango horario + precio con simbolo de moneda del negocio) para mostrar resumen tarifario en la lista detallada sin calculo en frontend.
  - `Sp_Combos_EspaciosPorNegocio` devuelve etiqueta de combo en formato: `Codigo - Nombre (Tipo suelo)`.
  - `Sp_UsuariosNegocio_ActualizarRol` devuelve error si no encuentra filas para el negocio.

### 33_Reservas_Historial_Recordatorio.sql
- SP nuevos:
  - `Sp_Reservas_Historial`
  - `Sp_Reservas_ObtenerParaRecordatorio`
- Objetivo:
  - soportar historial de acciones por reserva (drawer backend-driven)
  - habilitar recordatorio manual por seleccion (acciones por lote)

### 34_Clientes_NombreEquipo_Reservas.sql
- Alter tabla:
  - `Clientes.NombreEquipo` (NVARCHAR(120), NULL)
- SP actualizados:
  - `Sp_Combos_Clientes`
  - `Sp_Clientes_Listar`
  - `Sp_Clientes_ObtenerPorId`
  - `Sp_Clientes_Crear`
  - `Sp_Clientes_Actualizar`
  - `Sp_Reservas_Listar`
  - `Sp_Reservas_CalendarioEventos`
  - `Sp_Combos_ReservasPorNegocio`
  - `Sp_Reservas_ObtenerParaRecordatorio`
- Objetivo:
  - registrar nombre de equipo en maestro de clientes
  - mostrar el equipo en calendario y vistas de reserva (listado/detalle/edicion/acciones)

### 35_Maestros_FormasPago.sql
- Tabla:
  - `FormasPago`
- SP actualizados:
  - `Sp_Seguridad_SeedModulosPermisosBase` (agrega modulo `MAESTROS`)
  - `Sp_Pagos_Listar`
  - `Sp_Pagos_Crear`
  - `Sp_Pagos_Actualizar`
- SP nuevos:
  - `Sp_Combos_FormasPago`
  - `Sp_Maestros_Monedas_Listar`
  - `Sp_Maestros_Monedas_Crear`
  - `Sp_Maestros_Monedas_Actualizar`
  - `Sp_Maestros_Monedas_Eliminar`
  - `Sp_Maestros_TiposSuelo_Listar`
  - `Sp_Maestros_TiposSuelo_Crear`
  - `Sp_Maestros_TiposSuelo_Actualizar`
  - `Sp_Maestros_TiposSuelo_Eliminar`
  - `Sp_Maestros_TiposDeporte_Listar`
  - `Sp_Maestros_TiposDeporte_Crear`
  - `Sp_Maestros_TiposDeporte_Actualizar`
  - `Sp_Maestros_TiposDeporte_Eliminar`
  - `Sp_Maestros_FormasPago_Listar`
  - `Sp_Maestros_FormasPago_Crear`
  - `Sp_Maestros_FormasPago_Actualizar`
  - `Sp_Maestros_FormasPago_Eliminar`
- Objetivo:
  - habilitar pestana `Maestros` para mantenimiento backend-driven de catalogos base.
  - desacoplar formas de pago del valor fijo y manejarlo por tabla/combos.

### 36_Maestros_PorNegocio_MonedasSuper.sql
- Tabla:
  - `MonedasSuperMaestro` (supermaestro LATAM con codigo/simbolo)
- Alter tablas (auditoria + alcance por negocio):
  - `Monedas`
  - `TiposSuelo`
  - `TiposDeporte`
  - `FormasPago`
- SP redefinidos:
  - `Sp_Maestros_MonedasSuper_Listar`
  - `Sp_Combos_Monedas` (ahora por `@NegocioId`)
  - `Sp_Combos_TiposSuelo` (ahora por `@NegocioId`)
  - `Sp_Combos_TiposDeporte` (ahora por `@NegocioId`)
  - `Sp_Combos_FormasPago` (ahora por `@NegocioId`)
  - `Sp_Maestros_Monedas_*`, `Sp_Maestros_TiposSuelo_*`, `Sp_Maestros_TiposDeporte_*`, `Sp_Maestros_FormasPago_*` (todos por negocio)
  - `Sp_ConfiguracionClub_Actualizar` (valida moneda activa del negocio)
- Objetivo:
  - permitir que cada negocio maneje sus propios catálogos.
  - en `Maestros > Monedas`, registrar monedas del club seleccionando desde supermaestro.
  - `Sp_Maestros_Monedas_Crear` valida que por negocio solo se permita una moneda registrada.

### 37_Sedes_Ubicacion_Fotos.sql
- Alter tabla:
  - `Sedes.Latitud` (DECIMAL(10,7), NULL)
  - `Sedes.Longitud` (DECIMAL(10,7), NULL)
  - `Sedes.GooglePlaceId` (NVARCHAR(200), NULL)
  - `Sedes.GoogleMapsUrl` (NVARCHAR(500), NULL)
  - `Sedes.FotoPrincipalUrl` (NVARCHAR(500), NULL)
  - `Sedes.FotosUrlsCsv` (NVARCHAR(MAX), NULL)
- SP ajustados:
  - `Sp_Sedes_ObtenerPorId`
  - `Sp_Sedes_Crear`
  - `Sp_Sedes_Actualizar`
  - `Sp_Home_ListarSedesPublicas`
- Objetivo:
  - guardar ubicacion y fotos directamente en `Sedes` (sin tabla nueva)
  - exponer coordenadas/foto para mostrar mini mapa en el portal cliente.
  - soportar una foto principal y galeria alternativa por sede.

### 20260406_Maestros_TiposSueloSuperMaestro.sql
- Tabla nueva:
  - `TiposSueloSuperMaestro` (supermaestro de suelos con codigo y nombre).
- Alter tabla:
  - `TiposSuelo.TipoSueloSuperId` (FK hacia `TiposSueloSuperMaestro`).
  - indice unico `UX_TiposSuelo_Negocio_TipoSueloSuperId` para evitar duplicados por negocio.
- SP nuevos/ajustados:
  - `Sp_Maestros_TiposSueloSuper_Listar` (combo del supermaestro, solo nombre).
  - `Sp_Maestros_TiposSuelo_Crear` (alta por `@TipoSueloSuperId`).
  - `Sp_Maestros_TiposSuelo_Actualizar` (actualiza solo estado por negocio).
  - `Sp_Maestros_TiposSuelo_Listar` (incluye codigo/superId para UI).
- Objetivo:
  - en `Maestros > Tipos de suelo`, registrar tipos de suelo del club seleccionando desde supermaestro, igual que Monedas.

### 20260406_Maestros_TiposDeporteSuperMaestro.sql
- Tabla nueva:
  - `TiposDeporteSuperMaestro` (supermaestro de deportes con codigo y nombre).
- Alter tabla:
  - `TiposDeporte.TipoDeporteSuperId` (FK hacia `TiposDeporteSuperMaestro`).
  - indice unico `UX_TiposDeporte_Negocio_TipoDeporteSuperId` para evitar duplicados por negocio.
- SP nuevos/ajustados:
  - `Sp_Maestros_TiposDeporteSuper_Listar` (combo del supermaestro, solo nombre).
  - `Sp_Maestros_TiposDeporte_Crear` (alta por `@TipoDeporteSuperId`).
  - `Sp_Maestros_TiposDeporte_Actualizar` (actualiza solo estado por negocio).
  - `Sp_Maestros_TiposDeporte_Listar` (incluye codigo/superId para UI).
- Objetivo:
  - en `Maestros > Tipos de deporte`, registrar deportes del club seleccionando desde supermaestro, igual que Monedas.

### 19_Home_Whatsapp_Publico.sql (actualizacion 02/04/2026)
- Ajuste de contrato en `Sp_Home_ListarSedesPublicas`:
  - ahora devuelve `Sedes.FotosUrlsCsv` ademas de `FotoPrincipalUrl`.
- Objetivo:
  - mostrar galeria publica por sede (principal + alternativas) en modo cliente.

### 25_Sedes_Horario_Crear_Actualizar.sql (actualizacion 02/04/2026)
- Ajuste de validaciones en:
  - `Sp_Sedes_Crear`
  - `Sp_Sedes_Actualizar`
- Reglas nuevas de fotos por sede:
  - maximo 6 imagenes por sede (`1 principal + 5 alternativas`)
  - no permite fotos alternativas sin foto principal.
- Objetivo:
  - asegurar integridad del contrato de imagenes desde backend (ADO.NET + SP), sin depender de validaciones front-end.

### 20260404_Carga_Ubigeo_SUNAT.sql (nuevo 04/04/2026)
- Tablas nuevas (carpeta `Basededatos/SportCenter/Tablas`):
  - `dbo.UbigeoDepartamentos`
  - `dbo.UbigeoProvincias`
  - `dbo.UbigeoDistritos`
- Script de carga (carpeta `Basededatos/SportCenter/Script`):
  - `20260404_Carga_Ubigeo_SUNAT.sql`
- Fuente oficial:
  - anexo SUNAT RS `000103-2023/SUNAT` (`Departamento`, `Provincia`, `Distrito`, `Ubigeo`).
- Cobertura cargada:
  - `25` departamentos
  - `196` provincias
  - `1889` distritos
- Relacionamiento:
  - `UbigeoProvincias.CodigoDepartamento -> UbigeoDepartamentos.CodigoDepartamento`
  - `UbigeoDistritos.CodigoDepartamento -> UbigeoDepartamentos.CodigoDepartamento`
  - `UbigeoDistritos.CodigoProvincia -> UbigeoProvincias.CodigoProvincia`

### 20260404_Clientes_Configuracion_Ubigeo.sql (nuevo 04/04/2026)
- Ajustes de estructura:
  - `Clientes.CodigoUbigeo` (`CHAR(6)`, nullable, FK a `UbigeoDistritos`)
  - `Negocios.CodigoUbigeo` (`CHAR(6)`, nullable, FK a `UbigeoDistritos`)
- Regla funcional:
  - si se informa direccion fiscal, ubigeo (distrito) es obligatorio.
  - si no se informa direccion fiscal, `CodigoUbigeo` se limpia a `NULL`.

### StoreProcedure (individuales 04/04/2026)
- Nuevos SP en `Basededatos/SportCenter/StoreProcedure`:
  - `dbo.Sp_Ubigeo_Departamentos_Listar.StoredProcedure.sql`
  - `dbo.Sp_Ubigeo_Provincias_Listar.StoredProcedure.sql`
  - `dbo.Sp_Ubigeo_Distritos_Listar.StoredProcedure.sql`
  - `dbo.Sp_Ubigeo_ObtenerPorCodigo.StoredProcedure.sql`
- SP reemplazados (individuales):
  - `dbo.Sp_Clientes_ObtenerPorId.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Crear.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Obtener.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Actualizar.StoredProcedure.sql`

### 20260404_TiposDocumentoIdentidadSunat_Clientes_Configuracion.sql (nuevo 04/04/2026)
- Tabla nueva (carpeta `Basededatos/SportCenter/Tablas`):
  - `dbo.TiposDocumentoIdentidadSunat`
- Script de estructura/datos (carpeta `Basededatos/SportCenter/Script`):
  - `20260404_TiposDocumentoIdentidadSunat_Clientes_Configuracion.sql`
- Cobertura SUNAT cargada:
  - `0` (Doc. trib. no dom. sin RUC)
  - `1` (DNI)
  - `4` (Carnet de extranjeria)
  - `6` (RUC)
  - `7` (Pasaporte)
  - `A` (Cedula diplomatica)
- Reglas aplicadas:
  - `Clientes.TipoDocumento` y `Negocios.TipoDocumentoFiscal` migran a codigo SUNAT.
  - longitud de columnas alineada a SUNAT: `NVARCHAR(2)` en `Clientes.TipoDocumento` y `Negocios.TipoDocumentoFiscal`.
  - se crean FK a `TiposDocumentoIdentidadSunat.CodigoSunat` para integridad.
  - `Sp_Comprobantes_Crear` usa el codigo SUNAT real del cliente para `CodigoTipoDocumentoClienteSunat`.
  - Clientes y Configuracion cargan el combo desde `Sp_Combos_TiposDocumentoIdentidadSunat`.

### 99_SP_Finales.sql
- Script consolidado con la **ultima version efectiva** de cada `CREATE OR ALTER PROCEDURE`.
- Se genera automaticamente desde todos los `.sql` de la carpeta `deStoreProcedures_DbSportCenter`.
- Uso recomendado:
  - ejecutar `00..33` normalmente
  - ejecutar `99_SP_Finales.sql` al final para evitar sobreescritura accidental por orden.
- Generacion:
  - ejecutar `Generate-99_SP_Finales.ps1`
  - el script recorre la carpeta, detecta SP duplicados y conserva la version del archivo de mayor orden (ultimas capas funcionales).

### 20260409_Pagos_Listado_Reserva_Paginacion.sql
- SP actualizados:
  - `Sp_Pagos_Listar`
  - `Sp_Pagos_ObtenerPorId`
  - `Sp_Pagos_Actualizar`
  - `Sp_Pagos_Eliminar`
  - `Sp_Pagos_Crear`
- SP nuevos:
  - `Sp_Combos_Reservas_Buscar`
  - `Sp_Pagos_EliminarPorReserva`
- Objetivo:
  - listar pagos agrupados por reserva (una sola fila por reserva) con filtro y paginacion backend.
  - exponer en listado de pagos el `SaldoPendiente` por reserva y `MonedaSimbolo` para encabezado monetario (origen: `MonedasSuperMaestro.Simbolo` segun moneda del negocio).
  - editar pagos por reserva mostrando cabecera + detalle de pagos (2 resultsets), incluyendo fecha + horario de reserva.
  - registrar pago con busqueda incremental de reserva por texto (Enter) y resumen referencial de reserva seleccionada.
  - limitar edicion de pago existente a `Observacion`.
  - `Sp_Pagos_Crear` aplica validacion de politica del negocio (sin pago / adelanto minimo / pago total 100%), evita sobrepago, exige que el segundo pago sea igual al saldo restante y actualiza estado segun pago acumulado.
  - al eliminar pago:
    - si la reserva queda sin pagos -> estado `Cancelada`.
    - si mantiene pagos -> estado `Confirmada`.
  - al eliminar pagos desde listado por reserva:
    - elimina todos los pagos de la reserva.
    - deja la reserva en estado `Cancelada`.

### 20260412_Pagos_Referencia_Ultimo_Comprobante.sql
- SP actualizados:
  - `Sp_Pagos_Listar`
  - `Sp_Pagos_ObtenerPorId`
- Objetivo:
  - agregar columna `Referencia` al listado de pagos por reserva.
  - la referencia muestra el **ultimo comprobante principal activo generado** de la reserva (`01` factura, `03` boleta, `RI` recibo interno), usando prioridad por `Id` (ultimo registro creado).
  - si el ultimo comprobante principal esta anulado, no se muestra referencia.
  - para boleta/factura, si tiene notas relacionadas activas (`07` NC o `08` ND), no se muestra referencia.
  - al emitir un nuevo comprobante principal para la reserva, la referencia pasa a mostrar ese ultimo documento.
  - en edicion de pagos (`Sp_Pagos_ObtenerPorId`) se expone `TieneComprobanteActivo` y `ReferenciaComprobante` para bloquear alta/eliminacion de pagos cuando la reserva ya tiene comprobante emitido.

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
24. Ejecutar `24_Sedes_Horario_NoLaborable.sql`.
25. Ejecutar `25_Sedes_Horario_Crear_Actualizar.sql`.
26. Ejecutar `26_Reservas_Validacion_Horario_Sede.sql`.
27. Ejecutar `27_Calendario_No_Atencion_Sede.sql`.
28. Ejecutar `28_Espacios_Deporte_Suelo_Catalogos.sql`.
29. Ejecutar `29_Reservas_ValidarDisponibilidad_Modal.sql`.
30. Ejecutar `30_Configuracion_Club_Monedas.sql`.
31. Ejecutar `31_Espacios_Tarifas_Base.sql`.
32. Ejecutar `32_Usuarios_Sede_Restriccion_Filtros.sql`.
33. Ejecutar `33_Reservas_Historial_Recordatorio.sql`.
34. Ejecutar `34_Clientes_NombreEquipo_Reservas.sql`.
35. Ejecutar `35_Maestros_FormasPago.sql`.
36. Ejecutar `36_Maestros_PorNegocio_MonedasSuper.sql`.
37. Ejecutar `37_Sedes_Ubicacion_Fotos.sql`.
38. Ejecutar `EXEC dbo.Sp_Seguridad_SeedModulosPermisosBase;` una vez.
39. Ejecutar `99_SP_Finales.sql` como post-deploy para asegurar contrato final de SP.
40. Ejecutar `Basededatos/SportCenter/Tablas/dbo.UbigeoDepartamentos.Table.sql`.
41. Ejecutar `Basededatos/SportCenter/Tablas/dbo.UbigeoProvincias.Table.sql`.
42. Ejecutar `Basededatos/SportCenter/Tablas/dbo.UbigeoDistritos.Table.sql`.
43. Ejecutar `Basededatos/SportCenter/Script/20260404_Carga_Ubigeo_SUNAT.sql`.
44. Ejecutar `Basededatos/SportCenter/Script/20260404_Clientes_Configuracion_Ubigeo.sql` (solo estructura).
45. Ejecutar los SP individuales en `Basededatos/SportCenter/StoreProcedure` (nuevos y reemplazados de ubigeo/clientes/configuracion).
46. Ejecutar `Basededatos/SportCenter/Tablas/dbo.TiposDocumentoIdentidadSunat.Table.sql`.
47. Ejecutar `Basededatos/SportCenter/Script/20260404_TiposDocumentoIdentidadSunat_Clientes_Configuracion.sql`.
48. Ejecutar los SP individuales actualizados/nuevos:
  - `dbo.Sp_Combos_TiposDocumentoIdentidadSunat.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Crear.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Listar.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Obtener.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Comprobantes_Crear.StoredProcedure.sql`
  - `dbo.Sp_SolicitudesPublicas_ConvertirAReserva.StoredProcedure.sql`
49. Ejecutar `Basededatos/SportCenter/Script/20260406_Maestros_TiposSueloSuperMaestro.sql`.
50. Ejecutar los SP individuales actualizados/nuevos:
  - `dbo.Sp_Maestros_TiposSueloSuper_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposSuelo_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposSuelo_Crear.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposSuelo_Actualizar.StoredProcedure.sql`
51. Ejecutar `Basededatos/SportCenter/Script/20260406_Maestros_TiposDeporteSuperMaestro.sql`.
52. Ejecutar los SP individuales actualizados/nuevos:
  - `dbo.Sp_Maestros_TiposDeporteSuper_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDeporte_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDeporte_Crear.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDeporte_Actualizar.StoredProcedure.sql`
53. Ejecutar `Basededatos/SportCenter/Script/20260406_Negocios_PoliticaConfirmacionPago.sql` (solo estructura).
54. Ejecutar los SP individuales actualizados:
  - `dbo.Sp_ConfiguracionClub_Obtener.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Reservas_Crear.StoredProcedure.sql`
  - `dbo.Sp_Reservas_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Reservas_CambiarEstadoRapido.StoredProcedure.sql`
  - `dbo.Sp_SolicitudesPublicas_ConvertirAReserva.StoredProcedure.sql`
55. Ejecutar `Basededatos/SportCenter/Script/20260406_Sedes_ConsideracionesReserva.sql` (solo estructura).
56. Ejecutar los SP individuales actualizados:
  - `dbo.Sp_Sedes_Crear.StoredProcedure.sql`
  - `dbo.Sp_Sedes_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Sedes_ObtenerPorId.StoredProcedure.sql`
  - `dbo.Sp_Home_ListarSedesPublicas.StoredProcedure.sql`
  - `dbo.Sp_Home_BuscarEspaciosDisponibles.StoredProcedure.sql`
57. Ejecutar `Basededatos/SportCenter/Script/20260406_Clientes_Nombres_Apellidos.sql` (solo estructura).
58. Ejecutar los SP individuales actualizados:
  - `dbo.Sp_Clientes_Crear.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Clientes_ObtenerPorId.StoredProcedure.sql`
59. Ejecutar `Basededatos/SportCenter/Script/20260406_Combos_TiposDocumentoIdentidadSunat_Formato.sql`.
60. Ejecutar el SP actualizado:
  - `dbo.Sp_Combos_TiposDocumentoIdentidadSunat.StoredProcedure.sql`
61. Ejecutar `Basededatos/SportCenter/Script/20260406_Clientes_Listar_FiltroActivo.sql`.
62. Ejecutar el SP actualizado:
  - `dbo.Sp_Clientes_Listar.StoredProcedure.sql`
63. Ejecutar `Basededatos/SportCenter/Script/20260406_Sp_Clientes_Crear_NombresApellidos.sql` si al crear cliente aparece error de parametros en `Sp_Clientes_Crear`.
64. Ejecutar `Basededatos/SportCenter/Script/20260406_Clientes_NegocioId_SinTablaPuente.sql` (estructura y migracion de datos).
65. Ejecutar los SP individuales actualizados:
  - `dbo.Sp_Clientes_Crear.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Clientes_ObtenerPorId.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Listar.StoredProcedure.sql`
  - `dbo.Sp_Clientes_Eliminar.StoredProcedure.sql` (inactivacion logica)
  - `dbo.Sp_Combos_Clientes.StoredProcedure.sql`
  - `dbo.Sp_SolicitudesPublicas_ConvertirAReserva.StoredProcedure.sql`
66. Ejecutar scripts de validaciones operativas y listado:
  - `Basededatos/SportCenter/Script/20260407_Validaciones_Inactivacion_Clientes_Espacios_ReservasListado.sql`
  - `Basededatos/SportCenter/Script/20260407_Clientes_Documento_Reglas.sql`
  - `Basededatos/SportCenter/Script/20260407_Sp_Clientes_Actualizar_Validaciones.sql`
  - `Basededatos/SportCenter/Script/20260407_Sp_Clientes_Eliminar_Validaciones.sql`
  - `Basededatos/SportCenter/Script/20260407_Sp_Espacios_Actualizar_Validaciones.sql`
  - `Basededatos/SportCenter/Script/20260407_Sp_Espacios_Eliminar_Validaciones.sql`
  - `Basededatos/SportCenter/Script/20260407_Sp_Reservas_Listar_Saldo.sql`
67. Ejecutar `Basededatos/SportCenter/Script/20260407_Reservas_Pago_Fecha_NumOperacion.sql`.
68. Ejecutar los SP individuales actualizados:
  - `dbo.Sp_Reservas_Crear.StoredProcedure.sql`
  - `dbo.Sp_Reservas_Actualizar.StoredProcedure.sql`
69. Ejecutar `Basededatos/SportCenter/Script/20260407_Reservas_Cotizar_PoliticaSinTarifa.sql`.
70. Ejecutar `Basededatos/SportCenter/Script/20260407_Reservas_Listar_Cliente_Equipo_Columnas.sql`.
71. Ejecutar `Basededatos/SportCenter/Script/20260408_Sp_Reservas_CalendarioEventos_TotalReserva.sql`.
72. Ejecutar `Basededatos/SportCenter/Script/20260409_Pagos_Listado_Reserva_Paginacion.sql`.
73. Ejecutar `Basededatos/SportCenter/Script/20260409_DocumentosComprobante_Emision_Series.sql`.
74. Ejecutar/actualizar los SP individuales:
  - `dbo.Sp_ConfiguracionClub_Obtener.StoredProcedure.sql`
  - `dbo.Sp_ConfiguracionClub_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDocumentoComprobanteSuper_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDocumentoComprobante_Listar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDocumentoComprobante_Crear.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDocumentoComprobante_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Maestros_TiposDocumentoComprobante_Eliminar.StoredProcedure.sql`
  - `dbo.Sp_Combos_DocumentosComprobanteNegocio.StoredProcedure.sql`
  - `dbo.Sp_Configuracion_SeriesDocumentoComprobante_Listar.StoredProcedure.sql`
  - `dbo.Sp_Configuracion_SeriesDocumentoComprobante_Guardar.StoredProcedure.sql`
  - `dbo.Sp_Configuracion_SeriesDocumentoComprobante_Eliminar.StoredProcedure.sql`
  - `dbo.Sp_Combos_SeriesDocumentoComprobante.StoredProcedure.sql`
  - `dbo.Sp_Sedes_SeriesDocumentoComprobante_Listar.StoredProcedure.sql`
  - `dbo.Sp_Sedes_SeriesDocumentoComprobante_Guardar.StoredProcedure.sql`
  - `dbo.Sp_Reservas_Actualizar.StoredProcedure.sql`
  - `dbo.Sp_Reservas_CambiarEstadoRapido.StoredProcedure.sql`
  - `dbo.Sp_Reservas_Eliminar.StoredProcedure.sql`

## Observaciones funcionales
- CRUD de modulos internos ejecuta operaciones por SP.
- Validacion de permisos `PuedeVer/PuedeCrear/PuedeEditar/PuedeEliminar` se resuelve por SP de seguridad.
- Auditoria de acciones CRUD se registra con `Sp_Auditoria_Registrar`.
- Home publica permite registro de solicitudes de reserva con codigo de seguimiento.
- Notificacion por correo configurable via `EmailSettings` (SMTP).
- Modulo de promociones permite descuentos por sede/espacio con vigencia por fecha y franja horaria.
- Panel privado muestra KPIs avanzados para operacion diaria y seguimiento mensual.
- Reservas integra FullCalendar con vista semana/dia/mes, arrastre para mover horarios y bloqueos operativos por espacio.
- Reservas permite cambio rapido de estado (confirmada/pagada/cancelada/no asistio) desde tabla y calendario.
- Configuracion del club/negocio:
  - incluye politica para confirmar reservas por pago.
  - opciones por negocio: sin pago, adelanto minimo por porcentaje, o pago total (100%).
  - si la politica exige pago, backend bloquea confirmacion cuando no cumple (en cambio rapido, crear/editar reserva y convertir solicitud).
  - permite cargar un logo del club (JPG/PNG -> WebP en bucket), reemplazarlo o quitarlo, y usa `LogoUrl` del negocio para mostrarlo en la barra de menu admin.
- Reservas agrega operaciones avanzadas backend-driven:
  - historial por reserva desde bitacora
  - recordatorio manual por seleccion de reservas
  - resumen operativo diario con KPI y vista por espacios
- Clientes y reservas:
  - se incorpora `NombreEquipo` en maestro de clientes
  - para cliente con documento RUC, el formulario solicita razon social.
  - para documento distinto de RUC, el formulario solicita nombres y apellidos por separado.
  - `NombresORazonSocial` se mantiene como campo de compatibilidad y se llena automaticamente (concatenado para persona natural).
  - el nombre de equipo se refleja en combos y visualizacion de reservas (calendario, listado y detalle)
- el combo de tipo de documento (clientes/configuracion) muestra formato de etiqueta `Nombre (CodigoSunat)` sin codigo interno.
- el listado de clientes soporta filtro de estado (todos/activos/inactivos) con consulta backend.
- la paginacion y busqueda del listado de clientes se resuelven en SQL Server via `Sp_Clientes_Listar` con `@Buscar`, `@Pagina`, `@TamanoPagina` y `@TotalRegistros OUTPUT`.
- la relacion cliente-negocio ahora es directa por `Clientes.NegocioId`; `NegocioClientes` queda deprecada.
  - la accion del listado de clientes inactiva registro (`Activo = 0`) y no realiza eliminacion fisica.
  - no se permite inactivar cliente si tiene reservas activas futuras (pendiente/confirmada/pagada), mostrando detalle de reservas a cancelar.
  - tipo de documento `Doc. trib. no dom. sin RUC` no exige numero de documento.
  - el indice unico de clientes excluye `TipoDocumento = 0`, permitiendo multiples registros no domiciliados sin RUC sin colision por duplicado.
  - para los demas tipos se valida numero de documento obligatorio, maximo 11 digitos y solo numerico.
- Automatizacion en segundo plano:
  - envia recordatorios por correo antes de la hora de reserva
  - marca no-show automatico segun tolerancia configurada
- Configuracion por sede en formulario de Sedes:
  - activar/desactivar notificaciones
  - minutos de anticipacion del recordatorio
  - minutos de tolerancia para no-show automatico
  - correo de notificacion del negocio (copia oculta en recordatorios)
  - numero de WhatsApp de contacto para chatear/coordinar
  - consideraciones de reserva (texto libre) para publicar reglas y condiciones de atencion en el portal publico
- Portal publico:
  - boton de WhatsApp visible solo si la sede habilito chat y registro numero.
  - boton tambien visible en tarjetas de espacios disponibles para iniciar chat inmediato.
  - tarjetas de sedes muestran foto principal y carrusel con fotos alternativas cuando existen.
  - muestra consideraciones de la sede en tarjetas de sedes y en resultados de espacios disponibles.
- Carga de imagenes en Sedes:
  - formulario de `Nueva sede` y `Editar sede` permite subir archivos (`jpg/png/webp`) en lugar de pegar URLs.
  - carga integrada a almacenamiento objeto compatible S3 (Cloudflare R2) desde backend.
  - primera imagen cargada se registra como `FotoPrincipalUrl`; el resto va a `FotosUrlsCsv`.
  - limite tecnico y funcional de 6 imagenes por sede validado en front, controller y SP.
- Boton publico "Software para Clubes":
  - pide contrasena para crear la cuenta del dueno
  - valida codigo CAPTCHA
  - crea negocio/sede en el registro inicial
  - activa prueba automatica de 1 mes
- Control de suscripcion:
  - al vencer prueba o plan, el negocio queda bloqueado para operar modulos.
  - para reactivar se usa `Sp_NegociosSuscripcion_ActivarPlan`.
- Usuarios por sede:
  - si el rol es no administrador, la sede es obligatoria.
  - el backend restringe combos/listados/reportes/metricas a la sede asignada.
- Espacios deportivos:
  - no se permite cambiar a mantenimiento/inactivo ni inactivar espacio si el espacio tiene reservas activas futuras (pendiente/confirmada/pagada), mostrando detalle de reservas a cancelar.
- Reservas (listado general):
  - incorpora columnas de `Precio de Espacio` y `Saldo pendiente`.
  - agrega paginacion de 20 en 20 resuelta desde `Sp_Reservas_Listar` (no en memoria), manteniendo filtros de rango/sede/espacio/estados.
  - separa `Cliente` y `Equipo` en columnas distintas (ya no concatenadas en una sola celda).
  - importes de `Precio de Espacio` y `Saldo pendiente` se muestran alineados a la derecha.
- Pagos:
  - el listado principal se renderiza por reserva (una fila por reserva), no por cada movimiento de pago.
  - el listado usa paginacion backend (`20` por pagina) y filtro por texto.
  - el listado muestra columnas `Monto (<simbolo moneda>)` (monto total de la reserva) y `Saldo` por reserva.
  - la edicion permite solo actualizar `Observacion` de pagos existentes, agregar pago nuevo y/o marcar pagos para eliminacion con confirmacion previa.
  - en edicion de pagos se muestra tarjeta `Saldo` dentro de `Datos de la reserva`.
  - al eliminar todos los pagos de una reserva, la reserva queda en estado `Cancelada` y ya no aparece en el listado de pagos.
- Reservas (tarjetas del calendario):
  - ya no muestran nombre de espacio en eventos `RESERVA`.
  - muestran `Horario` y debajo `Precio de Espacio` usando `TotalReserva` expuesto por `Sp_Reservas_CalendarioEventos`.
- Reservas (pop-up crear/editar):
  - agrega `Comentario` en tabla `Reservas` y en formulario modal.
  - habilita cotizacion automatica por horario (`tarifa + promocion`) con endpoint `Sp_Reservas_Cotizar`.
  - actualizacion 13/04/2026: `Sp_Reservas_Cotizar` corrige mapeo de dia para domingo (`DiaSemana = 0`) y alinea la cotizacion con la configuracion de tarifas del modulo Espacios.
  - politica de pago del negocio visible en modal al crear/editar.
  - el estado no se elige al crear; se calcula segun pago registrado y politica del negocio.
  - soporte de registro de pago en creacion/edicion de reserva con forma de pago.
  - al marcar `Registrar pago`, se habilitan `Fecha de pago` (por defecto hoy, no permite fecha futura) y `N° Operacion` (opcional, solo alfanumerico).
  - en edicion, si el pago acumulado llega al 100% del precio del espacio, el backend ajusta automaticamente el estado a `Pagada`.
  - la politica de confirmacion se muestra siempre, incluso cuando no existe tarifa para el horario; en ese caso se permite ingreso manual del precio del espacio.
  - limite maximo de 2 pagos por reserva validado en `Sp_Reservas_Actualizar` y `Sp_Pagos_Crear`.
- Documentos de comprobante y emision:
  - se incorpora supermaestro `TiposDocumentoComprobanteSuperMaestro` con codigos SUNAT, campos `Tributario` y `Habilitado`, y documento interno `RI (Recibo Interno)`.
  - el seed de supermaestro carga el Catalogo SUNAT No. 01 completo (RS 245-2017/SUNAT, Anexo N. 8) y habilita por defecto `01-Factura` y `03-Boleta`.
  - se incorpora configuracion por negocio en `NegociosTiposDocumentoComprobante`.
  - se incorpora configuracion de series por negocio en `NegociosSeriesDocumentoComprobante`.
  - se incorpora configuracion opcional de serie por sede en `SedesSeriesDocumentoComprobante`.
  - `Negocios` agrega `EmisionComprobantesElectronicos`, `EmisionReciboInterno` y `PorcentajeIgv`.
  - en Configuracion se administra series por documento y activacion de emision tributaria/no tributaria.
  - en `Sedes > Documentos y series` se listan solo documentos activos en Maestros que ademas tienen al menos una serie activa configurada en Configuracion para ese mismo documento.
  - al agregar/inactivar series desde la misma seccion, tambien se persisten inmediatamente los checks `EmisionComprobantesElectronicos` y `EmisionReciboInterno` en `Negocios` mediante `Sp_ConfiguracionClub_ActualizarEmision`.
  - en Maestros se agrega mantenimiento de tipos de documento de comprobante por negocio.
- Cancelacion de reservas:
  - no se permite cancelar (editar, cambio rapido o boton cancelar) si la reserva tiene pagos registrados.
  - el sistema muestra validacion indicando eliminar primero los pagos.
- Comprobantes (emision desde reservas pagadas):
  - se agrega `Sp_Combos_ReservasPagadas_Buscar` para buscar reservas solo en estado `Pagada`.
  - `Sp_Combos_ReservasPagadas_Buscar` excluye reservas que ya tengan comprobante activo (`Estado <> Anulado`) para evitar doble emision.
  - `Sp_Comprobantes_Crear` valida reserva pagada, documento habilitado por negocio y serie habilitada por sede.
  - `Sp_Comprobantes_Crear` valida que la reserva pertenezca al negocio y bloquea duplicado de comprobante por reserva activa.
  - `Sp_Comprobantes_Crear` soporta documento `RI` (recibo interno), forzando `Igv = 0` y `SubTotal = Total`.
  - `Sp_Comprobantes_ObtenerPorId` devuelve `CodigoDocumentoComprobante` para la UI.
  - `Sp_Comprobantes_Listar` muestra etiquetas legibles para tipo (`Factura/Boleta/Recibo Interno`) y estado.
- Pagos (integracion con comprobantes):
  - el listado agrega botones `Emitir CPE` y `Emitir Recibo` segun checks de configuracion del negocio.
  - los botones abren la emision con reserva precargada.
  - los botones `Emitir CPE` y `Emitir Recibo` solo se muestran cuando la reserva esta pagada al 100% (sin saldo).
  - `Sp_Pagos_Listar` expone banderas `PagadaCompleta` y `TieneComprobanteActivo` para decidir habilitacion de emision desde backend.
  - si la reserva ya tiene comprobante activo (estado distinto de anulado), no se muestran botones de emision en listado de pagos.
- Comprobantes (registro y validaciones comerciales):
  - en `Emitir comprobante`, el campo `Numero` queda solo lectura y se genera automaticamente al grabar (correlativo por `TipoComprobante + Serie`).
  - si falta `Serie` al grabar, se muestra validacion inline en el combo (`NegocioSerieId`) ademas del resumen general.
  - `Sp_Comprobantes_Crear` actualiza datos editables del cliente desde el formulario de comprobante (`Correo`, `TipoDocumento`, `NumeroDocumento`, `DireccionFiscal`, `CodigoUbigeo`).
  - para `Boleta (03)`: valida `Total <= 700` y tipo de documento cliente en (`0`, `1`).
  - para `Factura (01)`: valida que el tipo de documento del cliente sea `RUC (6)`.
  - en emision se muestra primero bloque de datos editables del cliente (incluyendo ubigeo), y luego calculo de `SubTotal/IGV/Total`.
  - `Sp_Comprobantes_Listar` ahora soporta `@Buscar`, `@Pagina`, `@TamanoPagina` y devuelve `@TotalRegistros` para paginacion backend 20x20.
  - en `Editar comprobante`, la reserva queda fija y el importe queda solo lectura.
  - `Sp_Comprobantes_Actualizar` solo permite editar datos del cliente cuando el comprobante esta en estado `Pendiente`.
  - si el comprobante no esta pendiente, la UI y backend dejan el registro en solo lectura.
  - al grabar crear/editar comprobante, la aplicacion recarga el mismo registro (no regresa al listado).
  - en el listado de comprobantes se agrega accion `Ver` para visualizar el comprobante.
  - para documentos no tributarios (`Recibo Interno`) la visualizacion genera un PDF interno con cabecera + detalle (cantidad 1 por reserva).
  - para documentos tributarios (`Factura/Boleta`) la visualizacion redirige a la URL de descarga del proveedor, cuando la URL exista.
  - `Sp_Comprobantes_Listar` expone `EsTributario` y `UrlDescargaProveedor` para la UI.
  - se agrega `Sp_Comprobantes_ObtenerVisualizacion` para obtener datos de cabecera/detalle usados en la vista previa.
  - `Sp_Comprobantes_ObtenerVisualizacion` incluye ubigeo descriptivo (`Distrito/Provincia/Departamento`) de negocio y cliente para mostrarlo en la vista previa.
- Comprobantes (NC/ND desde comprobante aceptado SUNAT - 11/04/2026):
  - el listado de comprobantes agrega botones `Generar NC` y `Generar ND`; solo se habilitan cuando el comprobante origen es `Factura/Boleta` y esta en estado `Aceptado`.
  - para `Recibo Interno (RI)` no se muestran botones `Generar NC/ND`.
  - el listado agrega columna `Referencia`:
    - en NC/ND muestra el comprobante origen (Factura/Boleta).
    - en Factura/Boleta muestra las notas relacionadas (NC/ND) activas.
  - si un comprobante ya tiene NC/ND relacionadas activas, el listado desactiva `Anular` y `Generar NC/ND`.
  - se agrega flujo nuevo `CreateNota` (no reemplaza `Create`) para registrar Nota de Credito o Nota de Debito con:
    - documento de referencia (tipo/serie/numero del comprobante origen).
    - tipo de nota SUNAT obligatorio (combo segun `NC` o `ND`).
    - tipo de documento fijo (`07` para NC, `08` para ND).
  - se agrega tabla `dbo.TiposNotaComprobanteSunat` (maestro SUNAT de motivos NC/ND).
  - `dbo.ComprobantesElectronicos` agrega:
    - `ComprobanteReferenciaId` (FK a `ComprobantesElectronicos.Id`).
    - `TipoNota` (`07` para NC, `08` para ND).
    - `TipoNotaCodigoSunat`.
  - se agrega `Sp_Combos_TiposNotaComprobanteSunat` para poblar el combo de tipo de nota.
  - `Sp_Comprobantes_Crear` valida y crea NC/ND solo con comprobante referencia valido (Factura/Boleta aceptada SUNAT) y tipo de nota SUNAT activo.
  - al generar `NC (07)`, el comprobante referencia se mantiene activo; la reemision del comprobante principal se habilita por regla de negocio en `Sp_Comprobantes_Crear` cuando existe NC activa sobre el comprobante principal.
  - `Sp_Comprobantes_ObtenerPorId`, `Sp_Comprobantes_Listar` y `Sp_Comprobantes_ObtenerVisualizacion` ahora soportan codigos `07/08` y datos de nota/referencia.
  - `Sp_Comprobantes_Listar` obtiene tipo/codigo/referencias desde `NegociosTiposDocumentoComprobante` + `TiposDocumentoComprobanteSuperMaestro` (sin mapeo rigido por Id), manteniendo comportamiento multi-negocio.
  - en columnas `Referencia` (listado de comprobantes y pagos), el prefijo del documento usa `TiposDocumentoComprobanteSuperMaestro.Abreviatura` (con fallback a `Nombre`).
  - en pagos y combos de reservas para comprobantes, la reserva se considera disponible para reemision cuando el comprobante principal activo tiene una NC activa asociada.
  - se agrega script `20260413_Comprobantes_ReemisionPorNC_IndiceReserva.sql` para convertir `IX_ComprobantesElectronicos_ReservaId` a indice no unico (la restriccion de reemision por NC se controla en SP y no en filtro de indice).
  - el indice `IX_ComprobantesElectronicos_ReservaId` pasa a unico filtrado por `NegocioId + ReservaId` para comprobante principal activo (`Estado <> 5` y `ComprobanteReferenciaId IS NULL`), permitiendo NC/ND y derivados futuros (`ComprobanteReferenciaId IS NOT NULL`) sin depender de codigos de `TipoComprobante`.
- Reservas (listado general):
  - se agrega `Sp_Reservas_ListadoResumen` para KPI global del listado (Pendientes, Pagadas, Saldo total) con los mismos filtros del listado general y sin efecto de paginacion.
  - actualizacion 13/04/2026: `Sp_Reservas_ListadoResumen` excluye reservas canceladas (`Estado = 5`) del conteo de reservas activas KPI y del saldo total acumulado.
