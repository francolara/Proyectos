-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Cabecera de ventas por empresa con referencia al asiento contable generado.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Agrega subtotal, total exonerado, total inafecto e ICBPER interno para alinear la provision de ventas con compras.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Agrega la columna Saldo para controlar el importe pendiente del comprobante de venta.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Unifica el estado de las provisiones de venta a PROVISIONADO.
-- =============================================

IF OBJECT_ID(N'dbo.VEN_Venta', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VEN_Venta
    (
        IdVenta INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_VEN_Venta PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdCliente INT NOT NULL,
        IdConfiguracionContabilizacion INT NOT NULL,
        IdAsiento INT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        TipoComprobante VARCHAR(3) NOT NULL,
        Serie VARCHAR(10) NOT NULL,
        Numero VARCHAR(20) NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_VEN_Venta_TipoCambio DEFAULT (1),
        BaseImponible DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_BaseImponible DEFAULT (0),
        TotalExonerado DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_TotalExonerado DEFAULT (0),
        TotalInafecto DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_TotalInafecto DEFAULT (0),
        Icbper DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Icbper DEFAULT (0),
        Igv DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Igv DEFAULT (0),
        Isc DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Isc DEFAULT (0),
        OtrosTributos DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_OtrosTributos DEFAULT (0),
        Redondeo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Redondeo DEFAULT (0),
        ImporteTotal DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_ImporteTotal DEFAULT (0),
        Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Saldo DEFAULT (0),
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_VEN_Venta_Estado DEFAULT (N'PROVISIONADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_VEN_Venta_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_ADM_Cliente
            FOREIGN KEY (IdCliente) REFERENCES dbo.ADM_Cliente (IdCliente);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_CON_ConfiguracionContabilizacion
            FOREIGN KEY (IdConfiguracionContabilizacion) REFERENCES dbo.CON_ConfiguracionContabilizacion (IdConfiguracionContabilizacion);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT CK_VEN_Venta_Montos
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

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT UQ_VEN_Venta_Documento
            UNIQUE (IdEmpresa, IdCliente, TipoComprobante, Serie, Numero);
END;

IF COL_LENGTH(N'dbo.VEN_Venta', N'Saldo') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_Venta
        ADD Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Saldo DEFAULT (0);
END;
