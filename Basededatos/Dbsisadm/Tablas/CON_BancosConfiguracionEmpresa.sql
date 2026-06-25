-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea la configuracion de cuentas corrientes bancarias por empresa vinculada al plan de cuentas, titular y moneda operativa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_BancosConfiguracionEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_BancosConfiguracionEmpresa
    (
        IdBancoConfiguracionEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_BancosConfiguracionEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdBanco INT NOT NULL,
        NroCuentaCorriente VARCHAR(50) NOT NULL,
        Titular VARCHAR(200) NULL,
        IdMoneda INT NULL,
        IdPlanCuenta INT NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_BancosConfiguracionEmpresa_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_BancosConfiguracionEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_BancosConfiguracionEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_BancosConfiguracionEmpresa_CON_Bancos
            FOREIGN KEY (IdBanco) REFERENCES dbo.CON_Bancos (IdBanco);

    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_BancosConfiguracionEmpresa_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_BancosConfiguracionEmpresa_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT UQ_CON_BancosConfiguracionEmpresa_IdEmpresa_NroCuentaCorriente
            UNIQUE (IdEmpresa, NroCuentaCorriente);
END;

IF COL_LENGTH(N'dbo.CON_BancosConfiguracionEmpresa', N'Titular') IS NULL
BEGIN
    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD Titular VARCHAR(200) NULL;
END;

IF COL_LENGTH(N'dbo.CON_BancosConfiguracionEmpresa', N'IdMoneda') IS NULL
BEGIN
    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD IdMoneda INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_CON_BancosConfiguracionEmpresa_ADM_Moneda'
)
BEGIN
    ALTER TABLE dbo.CON_BancosConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_BancosConfiguracionEmpresa_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);
END;
