-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Firma de presupuesto privada en R2, concurrencia y auditoría.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'FirmaPresupuestoJson') IS NULL
    ALTER TABLE dbo.ConfiguracionHYT ADD FirmaPresupuestoJson NVARCHAR(2000) NULL;
GO
IF NOT EXISTS(SELECT 1 FROM sys.check_constraints WHERE parent_object_id=OBJECT_ID(N'dbo.ConfiguracionHYT') AND name=N'CK_ConfiguracionHYT_FirmaJSON')
    ALTER TABLE dbo.ConfiguracionHYT ADD CONSTRAINT CK_ConfiguracionHYT_FirmaJSON CHECK(FirmaPresupuestoJson IS NULL OR ISJSON(FirmaPresupuestoJson)=1);
GO
