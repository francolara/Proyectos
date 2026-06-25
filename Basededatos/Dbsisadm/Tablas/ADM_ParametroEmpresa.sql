-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Parametros administrativos y contables por empresa.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_ParametroEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_ParametroEmpresa
    (
        IdParametroEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_ParametroEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        TipoParametro VARCHAR(30) NOT NULL,
        CodigoParametro VARCHAR(100) NOT NULL,
        ValorParametro NVARCHAR(250) NOT NULL CONSTRAINT DF_ADM_ParametroEmpresa_ValorParametro DEFAULT (N''),
        DescripcionParametro NVARCHAR(300) NOT NULL CONSTRAINT DF_ADM_ParametroEmpresa_DescripcionParametro DEFAULT (N''),
        FecIni DATE NULL,
        FecFin DATE NULL,
        Activo BIT NOT NULL CONSTRAINT DF_ADM_ParametroEmpresa_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_ParametroEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_ParametroEmpresa
        ADD CONSTRAINT FK_ADM_ParametroEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.ADM_ParametroEmpresa
        ADD CONSTRAINT UQ_ADM_ParametroEmpresa_Empresa_Tipo_Codigo
            UNIQUE (IdEmpresa, TipoParametro, CodigoParametro);
END;
