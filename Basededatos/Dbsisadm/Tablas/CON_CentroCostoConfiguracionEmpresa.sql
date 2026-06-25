-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea la configuracion de centros de costo por empresa para uso operativo en asientos.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CentroCostoConfiguracionEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CentroCostoConfiguracionEmpresa
    (
        IdCentroCostoConfiguracionEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CentroCostoConfiguracionEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Codigo VARCHAR(20) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_CentroCostoConfiguracionEmpresa_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_CentroCostoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_CentroCostoConfiguracionEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_CentroCostoConfiguracionEmpresa
        ADD CONSTRAINT UQ_CON_CentroCostoConfiguracionEmpresa_IdEmpresa_Codigo
            UNIQUE (IdEmpresa, Codigo);
END;
