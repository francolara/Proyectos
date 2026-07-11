-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Registro de pagos y referencias de la suscripcion por cuenta administradora.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Prepara los cobros de suscripcion por cuenta para conciliacion y pasarela con metadatos externos y objetivos de aplicacion.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradoraSuscripcionPago', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
    (
        IdCuentaAdministradoraSuscripcionPago INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradoraSuscripcionPago PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        IdCuentaAdministradoraSuscripcion INT NOT NULL,
        IdCuentaAdministradoraSuscripcionMovimiento INT NULL,
        TipoPago NVARCHAR(30) NOT NULL,
        EstadoPago NVARCHAR(20) NOT NULL,
        Monto DECIMAL(12,2) NOT NULL,
        Moneda NVARCHAR(10) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_Moneda DEFAULT (N'PEN'),
        FechaPago DATETIME2(0) NOT NULL,
        FechaVencimiento DATE NULL,
        OperacionNumero NVARCHAR(100) NULL,
        EntidadFinanciera NVARCHAR(120) NULL,
        ReferenciaExterna NVARCHAR(120) NULL,
        ProveedorPasarela NVARCHAR(50) NULL,
        TransaccionPasarelaId NVARCHAR(120) NULL,
        PagoPasarelaId NVARCHAR(120) NULL,
        EstadoPasarela NVARCHAR(30) NULL,
        PayloadPasarela NVARCHAR(MAX) NULL,
        FechaConfirmacionPasarela DATETIME2(0) NULL,
        AccionAplicacion NVARCHAR(30) NULL,
        AplicarAlConfirmar BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_Aplicar DEFAULT (0),
        AplicadoSuscripcion BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_Aplicado DEFAULT (0),
        FechaAplicacion DATETIME2(0) NULL,
        UsuarioAplicacion NVARCHAR(450) NULL,
        TipoCobroObjetivo NVARCHAR(20) NULL,
        FechaInicioPlanObjetivo DATE NULL,
        DiasGraciaObjetivo INT NULL,
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        FechaActualizacion DATETIME2(0) NULL,
        UsuarioActualizacion NVARCHAR(450) NULL
    );

    CREATE NONCLUSTERED INDEX IX_SEG_CuentaAdministradoraSuscripcionPago_Cuenta_Fecha
        ON dbo.SEG_CuentaAdministradoraSuscripcionPago (IdCuentaAdministradora ASC, FechaPago DESC, IdCuentaAdministradoraSuscripcionPago DESC);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcionPago_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcionPago_SEG_CuentaAdministradoraSuscripcion
        FOREIGN KEY (IdCuentaAdministradoraSuscripcion) REFERENCES dbo.SEG_CuentaAdministradoraSuscripcion (IdCuentaAdministradoraSuscripcion);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionPago
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcionPago_SEG_CuentaAdministradoraSuscripcionMovimiento
        FOREIGN KEY (IdCuentaAdministradoraSuscripcionMovimiento) REFERENCES dbo.SEG_CuentaAdministradoraSuscripcionMovimiento (IdCuentaAdministradoraSuscripcionMovimiento);
END;
