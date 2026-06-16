-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Maestro de monedas del sistema.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_Moneda', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_Moneda
    (
        IdMoneda INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_Moneda PRIMARY KEY,
        CodigoMoneda VARCHAR(3) NOT NULL,
        NombreMoneda NVARCHAR(80) NOT NULL,
        SimboloMoneda NVARCHAR(10) NOT NULL,
        EsMonedaBase BIT NOT NULL CONSTRAINT DF_ADM_Moneda_EsMonedaBase DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_ADM_Moneda_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_Moneda_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_Moneda
        ADD CONSTRAINT UQ_ADM_Moneda_Codigo UNIQUE (CodigoMoneda);
END;
