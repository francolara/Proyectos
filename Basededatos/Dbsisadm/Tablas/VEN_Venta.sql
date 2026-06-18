-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Cabecera de ventas por empresa con referencia al asiento contable generado.
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
        Igv DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Igv DEFAULT (0),
        Isc DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Isc DEFAULT (0),
        OtrosTributos DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_OtrosTributos DEFAULT (0),
        Redondeo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Redondeo DEFAULT (0),
        ImporteTotal DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_ImporteTotal DEFAULT (0),
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_VEN_Venta_Estado DEFAULT (N'FACTURADO'),
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
                AND Igv >= 0
                AND Isc >= 0
                AND OtrosTributos >= 0
                AND Redondeo >= 0
                AND ImporteTotal >= 0
            );

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT UQ_VEN_Venta_Documento
            UNIQUE (IdEmpresa, IdCliente, TipoComprobante, Serie, Numero);
END;
