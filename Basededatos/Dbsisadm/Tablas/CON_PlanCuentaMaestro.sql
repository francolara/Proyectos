-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Plan de cuentas maestro interno base sin empresa, con ColBalance, moneda y tipo de cambio.
-- =============================================

IF OBJECT_ID(N'dbo.CON_PlanCuentaMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_PlanCuentaMaestro
    (
        IdPlanCuentaMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_PlanCuentaMaestro PRIMARY KEY,
        CodigoCuenta VARCHAR(20) NOT NULL,
        CodigoCuentaPadre VARCHAR(20) NULL,
        NombreCuenta NVARCHAR(200) NOT NULL,
        NivelCuenta TINYINT NOT NULL,
        ColBalance CHAR(1) NOT NULL,
        IdMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_IdMoneda DEFAULT (''),
        TipoCambio CHAR(1) NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_TipoCambio DEFAULT (''),
        AceptaMovimiento BIT NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_AceptaMovimiento DEFAULT (0),
        RequiereCentroCosto BIT NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_RequiereCentroCosto DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_Estado DEFAULT (1),
        Orden INT NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_Orden DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_PlanCuentaMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_PlanCuentaMaestro
        ADD CONSTRAINT UQ_CON_PlanCuentaMaestro_CodigoCuenta
            UNIQUE (CodigoCuenta);

END;
