-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Despliegue base del modulo Caja y Bancos con movimientos por cuenta corriente y detalle debe/haber.
-- =============================================

IF OBJECT_ID(N'dbo.BAN_MovimientoBanco', N'U') IS NULL
BEGIN
    PRINT 'La tabla dbo.BAN_MovimientoBanco debe desplegarse desde Tablas\\BAN_MovimientoBanco.sql';
END;

IF OBJECT_ID(N'dbo.BAN_MovimientoBancoDetalle', N'U') IS NULL
BEGIN
    PRINT 'La tabla dbo.BAN_MovimientoBancoDetalle debe desplegarse desde Tablas\\BAN_MovimientoBancoDetalle.sql';
END;
