-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Numeración interna de maestros sin modificar códigos históricos.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF NOT EXISTS(SELECT 1 FROM sys.sequences WHERE object_id=OBJECT_ID(N'dbo.SecuenciaCodigoInterno'))
    EXEC(N'CREATE SEQUENCE dbo.SecuenciaCodigoInterno AS BIGINT START WITH 1 INCREMENT BY 1 NO CYCLE;');
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.ServicioMaterial') AND name=N'UX_ServicioMaterial_CodigoInterno')
    CREATE UNIQUE INDEX UX_ServicioMaterial_CodigoInterno ON dbo.ServicioMaterial(CodigoInterno);
GO
