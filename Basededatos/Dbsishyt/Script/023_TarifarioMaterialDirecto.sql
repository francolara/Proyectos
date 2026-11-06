-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Identifica el soporte interno del tarifario directo por material, sin cambiar datos anteriores.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET QUOTED_IDENTIFIER ON;
IF COL_LENGTH(N'dbo.Tarifario',N'EsTarifarioMaterial') IS NULL
    ALTER TABLE dbo.Tarifario ADD EsTarifarioMaterial BIT NOT NULL CONSTRAINT DF_Tarifario_Material DEFAULT(0) WITH VALUES;
GO
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.Tarifario') AND name=N'UX_Tarifario_MaterialDirecto')
    CREATE UNIQUE INDEX UX_Tarifario_MaterialDirecto ON dbo.Tarifario(EsTarifarioMaterial) WHERE EsTarifarioMaterial=1;
GO
