-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Permisos base por modulo para cada rol de cuenta administradora.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_RolCuentaPermiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_RolCuentaPermiso
    (
        IdRolCuentaPermiso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_RolCuentaPermiso PRIMARY KEY,
        IdRolCuenta INT NOT NULL,
        IdModuloSistema INT NOT NULL,
        PuedeVer BIT NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_PuedeVer DEFAULT (0),
        PuedeCrear BIT NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_PuedeCrear DEFAULT (0),
        PuedeEditar BIT NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_PuedeEditar DEFAULT (0),
        PuedeEliminar BIT NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_PuedeEliminar DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_RolCuentaPermiso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_RolCuentaPermiso
        ADD CONSTRAINT UQ_SEG_RolCuentaPermiso UNIQUE (IdRolCuenta, IdModuloSistema);

    ALTER TABLE dbo.SEG_RolCuentaPermiso
        ADD CONSTRAINT FK_SEG_RolCuentaPermiso_SEG_RolCuenta
        FOREIGN KEY (IdRolCuenta) REFERENCES dbo.SEG_RolCuenta (IdRolCuenta);

    ALTER TABLE dbo.SEG_RolCuentaPermiso
        ADD CONSTRAINT FK_SEG_RolCuentaPermiso_SEG_ModuloSistema
        FOREIGN KEY (IdModuloSistema) REFERENCES dbo.SEG_ModuloSistema (IdModuloSistema);
END;
