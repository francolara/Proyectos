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
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega soporte para detracciones en compras con codigo SUNAT, porcentaje e importe descontado del saldo principal.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Agrega columnas para guardar fecha, estado y mensaje de validacion CPE en compras, ademas de la percepcion aplicada en cabecera.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Agrega soporte de retencion de renta de 4ta en compras con exoneracion, porcentaje e importe retenido en cabecera.

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
        ExoneracionRenta4ta BIT NOT NULL CONSTRAINT DF_COM_Compra_ExoneracionRenta4ta DEFAULT (0),
        PorcentajeRetencion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajeRetencion DEFAULT (0),
        Retencion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Retencion DEFAULT (0),
        TieneDetraccion BIT NOT NULL CONSTRAINT DF_COM_Compra_TieneDetraccion DEFAULT (0),
        IdDetraccionSunat INT NULL,
        PorcentajeDetraccion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajeDetraccion DEFAULT (0),
        ImporteDetraccion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImporteDetraccion DEFAULT (0),
        TienePercepcion BIT NOT NULL CONSTRAINT DF_COM_Compra_TienePercepcion DEFAULT (0),
        IdTipoPercepcion INT NULL,
        PorcentajePercepcion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajePercepcion DEFAULT (0),
        BasePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_BasePercepcion DEFAULT (0),
        ImportePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImportePercepcion DEFAULT (0),
        Observacion NVARCHAR(500) NULL,
        FechaValidacionCpe DATETIME2(0) NULL,
        EstadoValidacionCpe NVARCHAR(50) NULL,
        MensajeValidacionCpe NVARCHAR(500) NULL,
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
        ADD CONSTRAINT FK_COM_Compra_ADM_DetraccionSunat
            FOREIGN KEY (IdDetraccionSunat) REFERENCES dbo.ADM_DetraccionSunat (IdDetraccionSunat);

    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_TipoPercepcion
            FOREIGN KEY (IdTipoPercepcion) REFERENCES dbo.ADM_TipoPercepcion (IdTipoPercepcion);

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

IF COL_LENGTH(N'dbo.COM_Compra', N'ExoneracionRenta4ta') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD ExoneracionRenta4ta BIT NOT NULL CONSTRAINT DF_COM_Compra_ExoneracionRenta4ta DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'PorcentajeRetencion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD PorcentajeRetencion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajeRetencion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'Retencion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD Retencion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Retencion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'TieneDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TieneDetraccion BIT NOT NULL CONSTRAINT DF_COM_Compra_TieneDetraccion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'IdDetraccionSunat') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD IdDetraccionSunat INT NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'PorcentajeDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD PorcentajeDetraccion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajeDetraccion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'ImporteDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD ImporteDetraccion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImporteDetraccion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'FechaValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD FechaValidacionCpe DATETIME2(0) NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'TienePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TienePercepcion BIT NOT NULL CONSTRAINT DF_COM_Compra_TienePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'IdTipoPercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD IdTipoPercepcion INT NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'PorcentajePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD PorcentajePercepcion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'BasePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD BasePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_BasePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'ImportePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD ImportePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImportePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'EstadoValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD EstadoValidacionCpe NVARCHAR(50) NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'MensajeValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD MensajeValidacionCpe NVARCHAR(500) NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_COM_Compra_ADM_DetraccionSunat'
)
AND COL_LENGTH(N'dbo.COM_Compra', N'IdDetraccionSunat') IS NOT NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_DetraccionSunat
            FOREIGN KEY (IdDetraccionSunat) REFERENCES dbo.ADM_DetraccionSunat (IdDetraccionSunat);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_COM_Compra_ADM_TipoPercepcion'
)
AND COL_LENGTH(N'dbo.COM_Compra', N'IdTipoPercepcion') IS NOT NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_TipoPercepcion
            FOREIGN KEY (IdTipoPercepcion) REFERENCES dbo.ADM_TipoPercepcion (IdTipoPercepcion);
END;
