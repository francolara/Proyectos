-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Overrides de permisos por modulo a nivel de cuenta administradora para usuarios especificos.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_UsuarioCuentaPermiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_UsuarioCuentaPermiso
    (
        IdUsuarioCuentaPermiso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_UsuarioCuentaPermiso PRIMARY KEY,
        IdUsuarioCuentaAdministradora INT NOT NULL,
        IdModuloSistema INT NOT NULL,
        PuedeVer BIT NULL,
        PuedeCrear BIT NULL,
        PuedeEditar BIT NULL,
        PuedeEliminar BIT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaPermiso_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaPermiso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_UsuarioCuentaPermiso
        ADD CONSTRAINT UQ_SEG_UsuarioCuentaPermiso UNIQUE (IdUsuarioCuentaAdministradora, IdModuloSistema);

    ALTER TABLE dbo.SEG_UsuarioCuentaPermiso
        ADD CONSTRAINT FK_SEG_UsuarioCuentaPermiso_SEG_UsuarioCuentaAdministradora
        FOREIGN KEY (IdUsuarioCuentaAdministradora) REFERENCES dbo.SEG_UsuarioCuentaAdministradora (IdUsuarioCuentaAdministradora);

    ALTER TABLE dbo.SEG_UsuarioCuentaPermiso
        ADD CONSTRAINT FK_SEG_UsuarioCuentaPermiso_SEG_ModuloSistema
        FOREIGN KEY (IdModuloSistema) REFERENCES dbo.SEG_ModuloSistema (IdModuloSistema);
END;
