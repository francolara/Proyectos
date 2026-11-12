-- Firma: FRANCO LARA | Create date: 10/11/2026 | Habilita la ejecución de los procedimientos BCP desde HYT.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_TelecreditoHYT_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TelecreditoHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TelecreditoHYT_ConsultarLote TO hyt_app_executor;
GO
