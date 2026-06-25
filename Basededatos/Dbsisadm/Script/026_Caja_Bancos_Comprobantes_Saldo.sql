-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega el origen del comprobante y el importe aplicado en Caja y Bancos para descontar o restaurar saldo de compras y ventas.
-- =============================================

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'ModuloOperacionComprobante') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD ModuloOperacionComprobante CHAR(3) NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'IdRegistroComprobante') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD IdRegistroComprobante INT NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'ImporteAplicado') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD ImporteAplicado DECIMAL(18,2) NULL;
END;
