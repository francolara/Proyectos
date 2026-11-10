-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra permisos y acceso a proveedores, costos, pagos, trabajadores, vacaciones y asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'ASISTENCIA.CONSULTAR', N'Consultar jornadas y minutos extras aprobados'),
    (N'ASISTENCIA.EDITAR', N'Registrar jornadas y aprobar minutos extras'),
    (N'VACACIONES.CONSULTAR', N'Consultar los periodos y vacaciones de trabajadores'),
    (N'VACACIONES.EDITAR', N'Registrar periodos y gestionar vacaciones'),
    (N'TRABAJADORES.CONSULTAR', N'Consultar el maestro de trabajadores'),
    (N'TRABAJADORES.EDITAR', N'Crear y modificar trabajadores'),
    (N'TRABAJADORES.CUENTA', N'Consultar o cambiar la cuenta bancaria de trabajadores'),
    (N'PROVEEDORES.CONSULTAR', N'Consultar el maestro de proveedores'),
    (N'PROVEEDORES.EDITAR', N'Crear y modificar proveedores'),
    (N'PROVEEDORES.CUENTA', N'Consultar o cambiar la cuenta bancaria de proveedores'),
    (N'COSTOS_PROVEEDOR.CONSULTAR', N'Consultar costos de proveedores del presupuesto'),
    (N'COSTOS_PROVEEDOR.EDITAR', N'Registrar costos de proveedores del presupuesto'),
    (N'PAGOS_PROVEEDOR.CONSULTAR', N'Consultar saldos y pagos de proveedores'),
    (N'PAGOS_PROVEEDOR.EDITAR', N'Registrar pagos de proveedores')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CostoProveedor_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PagoProveedor_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CostoProveedor_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ProveedorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ProveedorHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PagoProveedor_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TrabajadorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TrabajadorHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_VacacionesHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PeriodoVacacionesHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_VacacionesHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_JornadaHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_JornadaHYT_Guardar TO hyt_app_executor;
GO
