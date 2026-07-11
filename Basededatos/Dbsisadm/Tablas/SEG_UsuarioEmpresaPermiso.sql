-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Overrides de permisos por modulo a nivel de empresa para usuarios especificos.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_UsuarioEmpresaPermiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_UsuarioEmpresaPermiso
    (
        IdUsuarioEmpresaPermiso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_UsuarioEmpresaPermiso PRIMARY KEY,
        IdUsuarioEmpresa INT NOT NULL,
        IdModuloSistema INT NOT NULL,
        PuedeVer BIT NULL,
        PuedeCrear BIT NULL,
        PuedeEditar BIT NULL,
        PuedeEliminar BIT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_UsuarioEmpresaPermiso_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_UsuarioEmpresaPermiso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_UsuarioEmpresaPermiso
        ADD CONSTRAINT UQ_SEG_UsuarioEmpresaPermiso UNIQUE (IdUsuarioEmpresa, IdModuloSistema);

    ALTER TABLE dbo.SEG_UsuarioEmpresaPermiso
        ADD CONSTRAINT FK_SEG_UsuarioEmpresaPermiso_SEG_UsuarioEmpresa
        FOREIGN KEY (IdUsuarioEmpresa) REFERENCES dbo.SEG_UsuarioEmpresa (IdUsuarioEmpresa);

    ALTER TABLE dbo.SEG_UsuarioEmpresaPermiso
        ADD CONSTRAINT FK_SEG_UsuarioEmpresaPermiso_SEG_ModuloSistema
        FOREIGN KEY (IdModuloSistema) REFERENCES dbo.SEG_ModuloSistema (IdModuloSistema);
END;
