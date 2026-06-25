-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Catalogo maestro interno de parametros base. No pertenece a una empresa.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_ParametroMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_ParametroMaestro
    (
        IdParametroMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_ParametroMaestro PRIMARY KEY,
        TipoParametro VARCHAR(30) NOT NULL,
        CodigoParametro VARCHAR(100) NOT NULL,
        ValorParametro NVARCHAR(250) NOT NULL CONSTRAINT DF_ADM_ParametroMaestro_ValorParametro DEFAULT (N''),
        DescripcionParametro NVARCHAR(300) NOT NULL CONSTRAINT DF_ADM_ParametroMaestro_DescripcionParametro DEFAULT (N''),
        FecIni DATE NULL,
        FecFin DATE NULL,
        Orden INT NOT NULL CONSTRAINT DF_ADM_ParametroMaestro_Orden DEFAULT (0),
        Activo BIT NOT NULL CONSTRAINT DF_ADM_ParametroMaestro_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_ParametroMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_ParametroMaestro
        ADD CONSTRAINT UQ_ADM_ParametroMaestro_Tipo_Codigo
            UNIQUE (TipoParametro, CodigoParametro);
END;
