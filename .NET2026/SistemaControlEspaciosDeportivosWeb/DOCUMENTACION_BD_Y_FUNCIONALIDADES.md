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
- `Sp_Sedes_Eliminar` y `Sp_Espacios_Eliminar` ahora retornan error cuando no existe el registro para el negocio.

### 04_Reservas_Pagos_Comprobantes.sql
- `Sp_Combos_EspaciosPorNegocio`
- `Sp_Combos_Clientes`
- `Sp_Combos_ReservasPorNegocio`
- `Sp_Reservas_Listar`
- `Sp_Reservas_ObtenerPorId`
- `Sp_Reservas_Crear`
- `Sp_Reservas_Actualizar`
- `Sp_Reservas_Eliminar`
- `Sp_Reservas_Eliminar` valida `@NegocioId` por join con `Sedes` y devuelve error si no encuentra la reserva.
- `Sp_Pagos_Listar`
- `Sp_Pagos_ObtenerPorId`
- `Sp_Pagos_Crear`
- `Sp_Pagos_Actualizar`
- `Sp_Pagos_Eliminar`
- `Sp_Pagos_Actualizar` y `Sp_Pagos_Eliminar` devuelven error si no existe el pago para el negocio (evita falso positivo en C#).
- `Sp_Comprobantes_Listar`
- `Sp_Comprobantes_ObtenerPorId`
- `Sp_Comprobantes_Crear`
- `Sp_Comprobantes_Actualizar`
- `Sp_Comprobantes_Eliminar`
- `Sp_Comprobantes_Actualizar` y `Sp_Comprobantes_Eliminar` devuelven error si no existe el comprobante para el negocio.

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
- `Sp_Pagos_Crear` (valida monto y evita sobrepago)
- `Sp_Pagos_Actualizar` (recalcula saldo en reserva nueva/anterior)
- `Sp_Pagos_Eliminar` (recalcula saldo de la reserva)
- `Sp_Pagos_Actualizar` y `Sp_Pagos_Eliminar` retornan error cuando el pago no existe para el negocio.

### 08_Reservas_Calendario_Filtros.sql
- `Sp_Reservas_Listar` (filtros por rango fecha, sede, espacio y estado simple `@Estado` o multiple `@EstadosCsv`)

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
- `Sp_Reservas_CambiarEstadoRapido` retorna error si la reserva no existe para el negocio.
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
  - `Sp_Panel_ObtenerMetricas`
  - `Sp_Promociones_Listar`
- Regla funcional:
  - usuarios no administradores trabajan con una sola sede asignada y los filtros/listados se restringen en backend.
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
- Reservas agrega operaciones avanzadas backend-driven:
  - historial por reserva desde bitacora
  - recordatorio manual por seleccion de reservas
  - resumen operativo diario con KPI y vista por espacios
- Clientes y reservas:
  - se incorpora `NombreEquipo` en maestro de clientes
  - el nombre de equipo se refleja en combos y visualizacion de reservas (calendario, listado y detalle)
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
  - tarjetas de sedes muestran foto principal y carrusel con fotos alternativas cuando existen.
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
