-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Amplia la suscripcion por cuenta con contrato comercial, gracia y metadatos de pasarela de pago.
-- =============================================

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcion', 'TipoCobro') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD TipoCobro NVARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcion', 'DiasGracia') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD DiasGracia INT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_DiasGracia DEFAULT (5);
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcion', 'FechaFinGracia') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD FechaFinGracia DATE NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcion', 'FechaActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD FechaActualizacion DATETIME2(0) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcion', 'UsuarioActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD UsuarioActualizacion NVARCHAR(450) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionMovimiento', 'TipoCobroAnterior') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD TipoCobroAnterior NVARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionMovimiento', 'TipoCobroNuevo') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD TipoCobroNuevo NVARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionMovimiento', 'DiasGracia') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD DiasGracia INT NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionMovimiento', 'DiasExtra') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD DiasExtra INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_SEG_CuentaAdministradoraSuscripcionMovimiento_Cuenta_Fecha'
      AND object_id = OBJECT_ID('dbo.SEG_CuentaAdministradoraSuscripcionMovimiento')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_SEG_CuentaAdministradoraSuscripcionMovimiento_Cuenta_Fecha
        ON dbo.SEG_CuentaAdministradoraSuscripcionMovimiento (IdCuentaAdministradora ASC, FechaRegistro DESC);
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'ProveedorPasarela') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD ProveedorPasarela NVARCHAR(50) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'TransaccionPasarelaId') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD TransaccionPasarelaId NVARCHAR(120) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'PagoPasarelaId') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD PagoPasarelaId NVARCHAR(120) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'EstadoPasarela') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD EstadoPasarela NVARCHAR(30) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'PayloadPasarela') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD PayloadPasarela NVARCHAR(MAX) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'FechaConfirmacionPasarela') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD FechaConfirmacionPasarela DATETIME2(0) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'TipoCobroObjetivo') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD TipoCobroObjetivo NVARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'FechaInicioPlanObjetivo') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD FechaInicioPlanObjetivo DATE NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'DiasGraciaObjetivo') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD DiasGraciaObjetivo INT NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'FechaActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD FechaActualizacion DATETIME2(0) NULL;
END;

IF COL_LENGTH('dbo.SEG_CuentaAdministradoraSuscripcionPago', 'UsuarioActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD UsuarioActualizacion NVARCHAR(450) NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_SEG_CuentaAdministradoraSuscripcionPago_Cuenta_Fecha'
      AND object_id = OBJECT_ID('dbo.SEG_CuentaAdministradoraSuscripcionPago')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_SEG_CuentaAdministradoraSuscripcionPago_Cuenta_Fecha
        ON dbo.SEG_CuentaAdministradoraSuscripcionPago (IdCuentaAdministradora ASC, FechaPago DESC, IdCuentaAdministradoraSuscripcionPago DESC);
END;
