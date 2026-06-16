-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Cuenta administradora o estudio contable titular de una o varias empresas.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradora', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradora
    (
        IdCuentaAdministradora INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradora PRIMARY KEY,
        CodigoCuenta VARCHAR(20) NOT NULL,
        NombreCuenta NVARCHAR(200) NOT NULL,
        CorreoPrincipal NVARCHAR(256) NOT NULL,
        TelefonoPrincipal NVARCHAR(30) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradora_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradora_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_CuentaAdministradora
        ADD CONSTRAINT UQ_SEG_CuentaAdministradora_CodigoCuenta UNIQUE (CodigoCuenta);
END;
