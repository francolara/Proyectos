-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Configuracion operativa principal de la cuenta administradora para opciones del panel General.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradoraConfiguracion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradoraConfiguracion
    (
        IdCuentaAdministradoraConfiguracion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradoraConfiguracion PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        NombreResponsablePrincipal NVARCHAR(180) NULL,
        CorreoAdministrativo NVARCHAR(256) NULL,
        TelefonoAdministrativo NVARCHAR(30) NULL,
        IdEmpresaPredeterminada INT NULL,
        ObservacionAdministrativa NVARCHAR(400) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraConfiguracion_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraConfiguracion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        FechaActualizacion DATETIME2(0) NULL,
        UsuarioActualizacion NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_CuentaAdministradoraConfiguracion
        ADD CONSTRAINT UQ_SEG_CuentaAdministradoraConfiguracion_IdCuentaAdministradora UNIQUE (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraConfiguracion
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraConfiguracion_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraConfiguracion
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraConfiguracion_SEG_Empresa
        FOREIGN KEY (IdEmpresaPredeterminada) REFERENCES dbo.SEG_Empresa (IdEmpresa);
END;
