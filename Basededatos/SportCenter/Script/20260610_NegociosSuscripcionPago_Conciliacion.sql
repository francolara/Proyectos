-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Agrega columnas de conciliacion y aplicacion automatica de cobros sobre NegociosSuscripcionPago.
-- =============================================
IF OBJECT_ID(N'dbo.NegociosSuscripcionPago', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'AccionAplicacion') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD AccionAplicacion NVARCHAR(30) NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'AplicarAlConfirmar') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD AplicarAlConfirmar BIT NOT NULL CONSTRAINT DF_NegociosSuscripcionPago_AplicarAlConfirmar DEFAULT ((0));

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'AplicadoSuscripcion') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD AplicadoSuscripcion BIT NOT NULL CONSTRAINT DF_NegociosSuscripcionPago_AplicadoSuscripcion DEFAULT ((0));

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'FechaAplicacion') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD FechaAplicacion DATETIME2(7) NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'UsuarioAplicacion') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD UsuarioAplicacion NVARCHAR(200) NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'TipoCobroObjetivo') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD TipoCobroObjetivo NVARCHAR(20) NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'FechaInicioPlanObjetivo') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD FechaInicioPlanObjetivo DATE NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcionPago', 'DiasGraciaObjetivo') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD DiasGraciaObjetivo INT NULL;
END
