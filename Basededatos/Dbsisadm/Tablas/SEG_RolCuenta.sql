-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Catalogo de roles base para la cuenta administradora.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_RolCuenta', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_RolCuenta
    (
        IdRolCuenta INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_RolCuenta PRIMARY KEY,
        CodigoRolCuenta VARCHAR(30) NOT NULL,
        NombreRolCuenta NVARCHAR(100) NOT NULL,
        DescripcionRol NVARCHAR(250) NULL,
        EsRolSistema BIT NOT NULL CONSTRAINT DF_SEG_RolCuenta_EsRolSistema DEFAULT (1),
        Estado BIT NOT NULL CONSTRAINT DF_SEG_RolCuenta_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_RolCuenta_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_RolCuenta
        ADD CONSTRAINT UQ_SEG_RolCuenta_CodigoRolCuenta UNIQUE (CodigoRolCuenta);
END;
