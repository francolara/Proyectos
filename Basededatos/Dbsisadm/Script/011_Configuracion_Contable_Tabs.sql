-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Crea configuracion contable por tabs: documentos, impuestos y provision.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Actualiza documentos con cuentas separadas y unifica configuracion de impuestos.
-- =============================================

IF OBJECT_ID(N'dbo.CON_TipoImpuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoImpuesto
    (
        IdTipoImpuesto INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoImpuesto PRIMARY KEY,
        CodigoSunat VARCHAR(10) NOT NULL,
        NombreImpuesto NVARCHAR(100) NOT NULL,
        IdPlanCuenta INT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_TipoImpuesto_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_TipoImpuesto
        ADD CONSTRAINT UQ_CON_TipoImpuesto_CodigoSunat UNIQUE (CodigoSunat);
END;

MERGE dbo.CON_TipoImpuesto AS destino
USING
(
    VALUES
        ('IGV', N'Impuesto General a las Ventas'),
        ('ISC', N'Impuesto Selectivo al Consumo'),
        ('IVAP', N'Impuesto a la Venta de Arroz Pilado'),
        ('ICBPER', N'Impuesto al Consumo de Bolsas Plasticas'),
        ('OTROS', N'Otros tributos')
) AS fuente (CodigoSunat, NombreImpuesto)
ON destino.CodigoSunat = fuente.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        NombreImpuesto = fuente.NombreImpuesto,
        Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, NombreImpuesto, Estado)
    VALUES (fuente.CodigoSunat, fuente.NombreImpuesto, 1);

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

    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_SEG_Empresa FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_ADM_TipoComprobante FOREIGN KEY (IdTipoComprobante) REFERENCES dbo.ADM_TipoComprobante (IdTipoComprobante);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaSoles FOREIGN KEY (IdCuentaVentaSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaDolares FOREIGN KEY (IdCuentaVentaDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraSoles FOREIGN KEY (IdCuentaCompraSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraDolares FOREIGN KEY (IdCuentaCompraDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD CONSTRAINT UQ_CON_DocumentoConfiguracionEmpresa UNIQUE (IdEmpresa, IdTipoComprobante);
END;

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

    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_SEG_Empresa FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_TipoImpuesto FOREIGN KEY (IdTipoImpuesto) REFERENCES dbo.CON_TipoImpuesto (IdTipoImpuesto);
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_PlanCuenta FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa ADD CONSTRAINT UQ_CON_TipoImpuestoConfiguracionEmpresa UNIQUE (IdEmpresa, IdTipoImpuesto);
END;

IF OBJECT_ID(N'dbo.CK_CON_ConfiguracionContabilizacion_EscenarioOperacion', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        DROP CONSTRAINT CK_CON_ConfiguracionContabilizacion_EscenarioOperacion;
END;

ALTER TABLE dbo.CON_ConfiguracionContabilizacion
    ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_EscenarioOperacion
        CHECK (EscenarioOperacion IN ('MERCADERIA', 'GASTO', 'SERVICIO', 'PROVISION'));
