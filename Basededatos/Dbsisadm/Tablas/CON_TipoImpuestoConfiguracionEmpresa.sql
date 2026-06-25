-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Configuracion unica de cuenta contable por impuesto y empresa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_TipoImpuestoConfiguracionEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
    (
        IdTipoImpuestoConfiguracionEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoImpuestoConfiguracionEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdTipoImpuesto INT NOT NULL,
        IdPlanCuenta INT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_TipoImpuestoConfiguracionEmpresa_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_TipoImpuestoConfiguracionEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_TipoImpuesto
            FOREIGN KEY (IdTipoImpuesto) REFERENCES dbo.CON_TipoImpuesto (IdTipoImpuesto);

    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT UQ_CON_TipoImpuestoConfiguracionEmpresa
            UNIQUE (IdEmpresa, IdTipoImpuesto);
END;
