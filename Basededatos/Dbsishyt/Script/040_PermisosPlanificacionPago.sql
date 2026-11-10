-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Concede acceso a planificacion con los permisos de pagos de personal.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_PlanPagoHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PlanPagoHYT_Guardar TO hyt_app_executor;
GO
