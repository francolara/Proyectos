-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Agrega saldo a compras y ventas y lo inicializa con el importe total actual.
-- =============================================

IF COL_LENGTH(N'dbo.COM_Compra', N'Saldo') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Saldo DEFAULT (0);
END;

UPDATE dbo.COM_Compra
SET Saldo = ImporteTotal
WHERE ISNULL(Saldo, 0) <> ImporteTotal;

IF COL_LENGTH(N'dbo.VEN_Venta', N'Saldo') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_Venta
        ADD Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Saldo DEFAULT (0);
END;

UPDATE dbo.VEN_Venta
SET Saldo = ImporteTotal
WHERE ISNULL(Saldo, 0) <> ImporteTotal;
