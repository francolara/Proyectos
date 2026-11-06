-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos limitados del marcado externo.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_Facturado_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Facturado_Obtener TO hyt_app_executor;
GO
