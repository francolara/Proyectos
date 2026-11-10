-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Autoriza consulta y edicion de obligaciones y polizas SCTR.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'OBLIGACIONES.CONSULTAR', N'Consultar AFP, ONP y polizas SCTR'),
    (N'OBLIGACIONES.EDITAR', N'Guardar AFP, ONP y polizas SCTR')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ObligacionesHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ObligacionPersonalHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PolizaSctrHYT_Guardar TO hyt_app_executor;
GO
