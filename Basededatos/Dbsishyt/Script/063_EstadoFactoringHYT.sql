-- Firma: FRANCO LARA | Create date: 14/11/2026 | Agrega el estado a los lotes nuevos y existentes.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.FactoringLoteHYT', N'Estado') IS NULL
    ALTER TABLE dbo.FactoringLoteHYT
        ADD Estado NVARCHAR(12) NOT NULL CONSTRAINT DF_FactoringLoteHYT_Estado DEFAULT N'REGISTRADO'
        CONSTRAINT CK_FactoringLoteHYT_Estado CHECK (Estado IN (N'REGISTRADO',N'ACEPTADO',N'CERRADO'));
GO
