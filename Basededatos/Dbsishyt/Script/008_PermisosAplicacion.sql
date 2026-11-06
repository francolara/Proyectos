-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Prepara rol SQL de ejecucion limitada sin crear login ni usuario.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF DATABASE_PRINCIPAL_ID(N'hyt_app_executor') IS NULL
    EXEC(N'CREATE ROLE [hyt_app_executor] AUTHORIZATION [dbo];');
DENY INSERT, UPDATE, DELETE ON SCHEMA::dbo TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.Moneda TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TipoDocumento TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TipoIGV TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TasaIGV TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ListarClientes TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_RegistrarTasaIGV TO hyt_app_executor;
-- No otorgar acceso directo a spu_RegistrarAuditoria ni a datos privados.
-- Los SP de operaciones futuras se concederan individualmente al implementarse.
GO
