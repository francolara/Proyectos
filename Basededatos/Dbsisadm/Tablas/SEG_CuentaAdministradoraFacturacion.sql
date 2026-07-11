-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Datos de facturacion por cuenta administradora para boleta, factura y contacto administrativo.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradoraFacturacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradoraFacturacion
    (
        IdCuentaAdministradoraFacturacion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradoraFacturacion PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        TipoComprobantePreferido VARCHAR(20) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraFacturacion_TipoComprobantePreferido DEFAULT ('BOLETA'),
        TipoDocumentoFacturacion VARCHAR(20) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraFacturacion_TipoDocumentoFacturacion DEFAULT ('DNI'),
        NumeroDocumento VARCHAR(20) NULL,
        NombreFacturacion NVARCHAR(200) NULL,
        RazonSocialFacturacion NVARCHAR(200) NULL,
        CorreoFacturacion NVARCHAR(256) NULL,
        TelefonoFacturacion NVARCHAR(30) NULL,
        DireccionFiscal NVARCHAR(250) NULL,
        Ubigeo VARCHAR(6) NULL,
        Distrito NVARCHAR(100) NULL,
        Provincia NVARCHAR(100) NULL,
        Departamento NVARCHAR(100) NULL,
        ObservacionFacturacion NVARCHAR(400) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraFacturacion_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraFacturacion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        FechaActualizacion DATETIME2(0) NULL,
        UsuarioActualizacion NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_CuentaAdministradoraFacturacion
        ADD CONSTRAINT UQ_SEG_CuentaAdministradoraFacturacion_IdCuentaAdministradora UNIQUE (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraFacturacion
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraFacturacion_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraFacturacion
        ADD CONSTRAINT CK_SEG_CuentaAdministradoraFacturacion_TipoComprobante
        CHECK (TipoComprobantePreferido IN ('BOLETA', 'FACTURA'));

    ALTER TABLE dbo.SEG_CuentaAdministradoraFacturacion
        ADD CONSTRAINT CK_SEG_CuentaAdministradoraFacturacion_TipoDocumento
        CHECK (TipoDocumentoFacturacion IN ('DNI', 'RUC', 'CE', 'PASAPORTE', 'OTRO'));
END;
