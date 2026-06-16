-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Plan de cuentas contable por empresa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_PlanCuenta', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_PlanCuenta
    (
        IdPlanCuenta INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_PlanCuenta PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdPlanCuentaPadre INT NULL,
        CodigoCuenta VARCHAR(20) NOT NULL,
        NombreCuenta NVARCHAR(200) NOT NULL,
        NivelCuenta TINYINT NOT NULL,
        NaturalezaSaldo CHAR(1) NOT NULL,
        AceptaMovimiento BIT NOT NULL CONSTRAINT DF_CON_PlanCuenta_AceptaMovimiento DEFAULT (0),
        RequiereCentroCosto BIT NOT NULL CONSTRAINT DF_CON_PlanCuenta_RequiereCentroCosto DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_CON_PlanCuenta_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_PlanCuenta_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_PlanCuenta
        ADD CONSTRAINT FK_CON_PlanCuenta_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_PlanCuenta
        ADD CONSTRAINT FK_CON_PlanCuenta_Padre
            FOREIGN KEY (IdPlanCuentaPadre) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_PlanCuenta
        ADD CONSTRAINT CK_CON_PlanCuenta_NaturalezaSaldo
            CHECK (NaturalezaSaldo IN ('D', 'H'));

    ALTER TABLE dbo.CON_PlanCuenta
        ADD CONSTRAINT CK_CON_PlanCuenta_NivelCuenta
            CHECK (NivelCuenta >= 1);

    ALTER TABLE dbo.CON_PlanCuenta
        ADD CONSTRAINT UQ_CON_PlanCuenta_IdEmpresa_CodigoCuenta
            UNIQUE (IdEmpresa, CodigoCuenta);
END;
