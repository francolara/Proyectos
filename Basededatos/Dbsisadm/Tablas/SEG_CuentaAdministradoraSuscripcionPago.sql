-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Registro de pagos y referencias de la suscripcion por cuenta administradora.
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
        AccionAplicacion NVARCHAR(30) NULL,
        AplicarAlConfirmar BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_Aplicar DEFAULT (0),
        AplicadoSuscripcion BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_Aplicado DEFAULT (0),
        FechaAplicacion DATETIME2(0) NULL,
        UsuarioAplicacion NVARCHAR(450) NULL,
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionPago_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

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
