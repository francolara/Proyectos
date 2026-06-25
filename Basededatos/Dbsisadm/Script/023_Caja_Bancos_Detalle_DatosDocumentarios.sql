-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Agrega referencias documentarias por linea al detalle de movimientos de caja y bancos.
-- =============================================

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'NumeroDocumento') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD NumeroDocumento VARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'TipoDocumento') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD TipoDocumento NVARCHAR(150) NULL;
END;

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'Serie') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD Serie VARCHAR(10) NULL;
END;

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'ReferenciaLinea') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD ReferenciaLinea NVARCHAR(100) NULL;
END;

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'TipoCambioLinea') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD TipoCambioLinea DECIMAL(18, 6) NULL;
END;
