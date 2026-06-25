-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Agrega titular y moneda a la configuracion de cuentas corrientes por empresa.
-- =============================================

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
