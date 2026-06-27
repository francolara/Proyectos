-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega importes por moneda en detalle de movimientos bancarios y asientos para guardar equivalencias en soles y dolares.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Agrega columnas TotalImporteS y TotalImporteD al detalle bancario y contable existente.

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'TotalImporteS') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteS DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'TotalImporteD') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteD DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'TotalImporteS') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteS DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'TotalImporteD') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteD DEFAULT (0);
END;
