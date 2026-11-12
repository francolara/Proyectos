-- Firma: FRANCO LARA | Create date: 10/11/2026 | Facturas vinculadas con un solo lote de factoring.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.FactoringFacturaHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FactoringFacturaHYT
    (
        ComprobanteId BIGINT NOT NULL CONSTRAINT PK_FactoringFacturaHYT PRIMARY KEY REFERENCES dbo.ComprobanteHYT(ComprobanteId),
        LoteId BIGINT NOT NULL REFERENCES dbo.FactoringLoteHYT(LoteId)
    );
END;
GO
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.FactoringFacturaHYT') AND name=N'IX_FactoringFacturaHYT_Lote')
    CREATE INDEX IX_FactoringFacturaHYT_Lote ON dbo.FactoringFacturaHYT(LoteId);
GO
