-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Cabecera de provisiones de compra por empresa con referencia al asiento contable generado.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Agrega totales exonerado e inafecto en cabecera de compra y mantiene ICBPER interno en cero.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Agrega la columna Saldo para controlar el importe pendiente del comprobante de compra.
-- =============================================

IF OBJECT_ID(N'dbo.COM_Compra', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.COM_Compra
    (
        IdCompra INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_COM_Compra PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdProveedor INT NOT NULL,
        IdConfiguracionContabilizacion INT NOT NULL,
        IdAsiento INT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        TipoComprobante VARCHAR(3) NOT NULL,
        Serie VARCHAR(10) NOT NULL,
        Numero VARCHAR(20) NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_COM_Compra_TipoCambio DEFAULT (1),
        BaseImponible DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_BaseImponible DEFAULT (0),
        TotalExonerado DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_TotalExonerado DEFAULT (0),
        TotalInafecto DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_TotalInafecto DEFAULT (0),
        Icbper DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Icbper DEFAULT (0),
        Igv DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Igv DEFAULT (0),
        Isc DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Isc DEFAULT (0),
        OtrosTributos DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_OtrosTributos DEFAULT (0),
        Redondeo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Redondeo DEFAULT (0),
        ImporteTotal DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImporteTotal DEFAULT (0),
        Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Saldo DEFAULT (0),
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_COM_Compra_Estado DEFAULT (N'PROVISIONADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_COM_Compra_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_Proveedor
            FOREIGN KEY (IdProveedor) REFERENCES dbo.ADM_Proveedor (IdProveedor);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_CON_ConfiguracionContabilizacion
            FOREIGN KEY (IdConfiguracionContabilizacion) REFERENCES dbo.CON_ConfiguracionContabilizacion (IdConfiguracionContabilizacion);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT CK_COM_Compra_Montos
            CHECK (
                BaseImponible >= 0
                AND TotalExonerado >= 0
                AND TotalInafecto >= 0
                AND Icbper >= 0
                AND Igv >= 0
                AND Isc >= 0
                AND OtrosTributos >= 0
                AND Redondeo >= 0
                AND Saldo >= 0
                AND ImporteTotal >= 0
            );

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT UQ_COM_Compra_Documento
            UNIQUE (IdEmpresa, IdProveedor, TipoComprobante, Serie, Numero);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'TotalExonerado') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TotalExonerado DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_TotalExonerado DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'TotalInafecto') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TotalInafecto DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_TotalInafecto DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'Icbper') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD Icbper DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Icbper DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'Saldo') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Saldo DEFAULT (0);
END;
