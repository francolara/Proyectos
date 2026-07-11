-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Catalogo de modulos del sistema con alcance de cuenta o empresa para resolver permisos por opcion.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_ModuloSistema', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_ModuloSistema
    (
        IdModuloSistema INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_ModuloSistema PRIMARY KEY,
        CodigoModulo VARCHAR(50) NOT NULL,
        NombreModulo NVARCHAR(150) NOT NULL,
        DescripcionModulo NVARCHAR(250) NULL,
        AlcanceModulo VARCHAR(20) NOT NULL,
        GrupoMenu NVARCHAR(100) NULL,
        OrdenMenu INT NOT NULL CONSTRAINT DF_SEG_ModuloSistema_OrdenMenu DEFAULT (0),
        EsVisibleMenu BIT NOT NULL CONSTRAINT DF_SEG_ModuloSistema_EsVisibleMenu DEFAULT (1),
        Estado BIT NOT NULL CONSTRAINT DF_SEG_ModuloSistema_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_ModuloSistema_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_ModuloSistema
        ADD CONSTRAINT UQ_SEG_ModuloSistema_CodigoModulo UNIQUE (CodigoModulo);

    ALTER TABLE dbo.SEG_ModuloSistema
        ADD CONSTRAINT CK_SEG_ModuloSistema_AlcanceModulo
        CHECK (AlcanceModulo IN ('CUENTA', 'EMPRESA'));
END;
