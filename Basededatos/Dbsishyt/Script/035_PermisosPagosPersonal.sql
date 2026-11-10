-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra permisos de estado de cuenta, movimientos, cierres, ceses y reportes de planillas.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'PAGOS_PERSONAL.CONSULTAR', N'Consultar el estado de cuenta del personal'),
    (N'PAGOS_PERSONAL.EDITAR', N'Registrar movimientos de personal'),
    (N'PAGOS_PERSONAL.CERRAR', N'Cerrar el mes de movimientos de personal'),
    (N'PAGOS_PERSONAL.CESES', N'Registrar ceses de personal'),
    (N'REPORTES_PLANILLAS.CONSULTAR', N'Consultar y exportar reportes de planillas')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_MovimientoPersonalHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CierrePersonalHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_EstadoCuentaPersonalHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CeseTrabajadorHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ReportePlanillasHYT_Consultar TO hyt_app_executor;
GO
