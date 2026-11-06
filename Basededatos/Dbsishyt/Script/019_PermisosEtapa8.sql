-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos limitados de dashboard y reportes.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_Dashboard_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Reporte_Listar TO hyt_app_executor;
GO
