-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Configuracion de cuentas contables por documento y empresa.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Separa cuentas por documento para compras y ventas en moneda soles/dolares.
-- =============================================

IF OBJECT_ID(N'dbo.CON_DocumentoConfiguracionEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_DocumentoConfiguracionEmpresa
    (
        IdDocumentoConfiguracionEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_DocumentoConfiguracionEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdTipoComprobante INT NOT NULL,
        IdCuentaVentaSoles INT NULL,
        IdCuentaVentaDolares INT NULL,
        IdCuentaCompraSoles INT NULL,
        IdCuentaCompraDolares INT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_DocumentoConfiguracionEmpresa_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_DocumentoConfiguracionEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_ADM_TipoComprobante
            FOREIGN KEY (IdTipoComprobante) REFERENCES dbo.ADM_TipoComprobante (IdTipoComprobante);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaSoles
            FOREIGN KEY (IdCuentaVentaSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaDolares
            FOREIGN KEY (IdCuentaVentaDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraSoles
            FOREIGN KEY (IdCuentaCompraSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraDolares
            FOREIGN KEY (IdCuentaCompraDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT UQ_CON_DocumentoConfiguracionEmpresa
            UNIQUE (IdEmpresa, IdTipoComprobante);
END;
