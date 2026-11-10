-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Habilita consulta de catálogos de referencia.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CatalogosPlanillasHYT_Consultar TO hyt_app_executor;
GO
