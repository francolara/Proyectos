-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Renombra NaturalezaSaldo a ColBalance y agrega IdMoneda/TipoCambio en plan de cuentas y maestro.
-- =============================================

IF OBJECT_ID(N'dbo.CON_PlanCuenta', N'U') IS NOT NULL
BEGIN
    IF OBJECT_ID(N'dbo.CK_CON_PlanCuenta_NaturalezaSaldo', N'C') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuenta DROP CONSTRAINT CK_CON_PlanCuenta_NaturalezaSaldo;
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuenta', N'NaturalezaSaldo') IS NOT NULL
       AND COL_LENGTH(N'dbo.CON_PlanCuenta', N'ColBalance') IS NULL
    BEGIN
        EXEC sp_rename N'dbo.CON_PlanCuenta.NaturalezaSaldo', N'ColBalance', N'COLUMN';
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuenta', N'ColBalance') IS NOT NULL
    BEGIN
        UPDATE dbo.CON_PlanCuenta
        SET ColBalance = CASE
                            WHEN ColBalance = 'D' THEN 'I'
                            WHEN ColBalance = 'H' THEN 'R'
                            ELSE ColBalance
                         END;
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuenta', N'IdMoneda') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuenta
            ADD IdMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_PlanCuenta_IdMoneda DEFAULT ('');
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuenta', N'TipoCambio') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuenta
            ADD TipoCambio CHAR(1) NOT NULL CONSTRAINT DF_CON_PlanCuenta_TipoCambio DEFAULT ('');
    END;
END;

IF OBJECT_ID(N'dbo.CON_PlanCuentaMaestro', N'U') IS NOT NULL
BEGIN
    IF OBJECT_ID(N'dbo.CK_CON_PlanCuentaMaestro_NaturalezaSaldo', N'C') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuentaMaestro DROP CONSTRAINT CK_CON_PlanCuentaMaestro_NaturalezaSaldo;
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuentaMaestro', N'NaturalezaSaldo') IS NOT NULL
       AND COL_LENGTH(N'dbo.CON_PlanCuentaMaestro', N'ColBalance') IS NULL
    BEGIN
        EXEC sp_rename N'dbo.CON_PlanCuentaMaestro.NaturalezaSaldo', N'ColBalance', N'COLUMN';
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuentaMaestro', N'ColBalance') IS NOT NULL
    BEGIN
        UPDATE dbo.CON_PlanCuentaMaestro
        SET ColBalance = CASE
                            WHEN ColBalance = 'D' THEN 'I'
                            WHEN ColBalance = 'H' THEN 'R'
                            ELSE ColBalance
                         END;
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuentaMaestro', N'IdMoneda') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuentaMaestro
            ADD IdMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_IdMoneda DEFAULT ('');
    END;

    IF COL_LENGTH(N'dbo.CON_PlanCuentaMaestro', N'TipoCambio') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_PlanCuentaMaestro
            ADD TipoCambio CHAR(1) NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_TipoCambio DEFAULT ('');
    END;
END;
