-- Firma: FRANCO LARA | Create date: 10/11/2026 | Identifica la asignación familiar aunque cambie su nombre.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ConceptoPlanillaHYT',N'CodigoSistema') IS NULL
    ALTER TABLE dbo.ConceptoPlanillaHYT ADD CodigoSistema VARCHAR(30) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.ConceptoPlanillaHYT') AND name=N'UX_ConceptoPlanillaHYT_CodigoSistema')
    CREATE UNIQUE INDEX UX_ConceptoPlanillaHYT_CodigoSistema ON dbo.ConceptoPlanillaHYT(CodigoSistema) WHERE CodigoSistema IS NOT NULL;
GO
IF NOT EXISTS (SELECT 1 FROM dbo.ConceptoPlanillaHYT WHERE CodigoSistema='ASIGNACION_FAMILIAR')
    UPDATE dbo.ConceptoPlanillaHYT SET CodigoSistema='ASIGNACION_FAMILIAR' WHERE Nombre=N'Asignación familiar';
GO
