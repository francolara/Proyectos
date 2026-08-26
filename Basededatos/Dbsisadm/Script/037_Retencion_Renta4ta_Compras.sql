-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Agrega retencion de renta de 4ta en compras por recibos por honorarios y crea el documento pendiente independiente para su pago.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Registra la cuenta maestra de R4TA como CodigoCuenta VARCHAR en lugar de IdPlanCuenta.

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

IF OBJECT_ID(N'dbo.COM_CompraRetencion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.COM_CompraRetencion
    (
        IdCompraRetencion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_COM_CompraRetencion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdCompra INT NOT NULL,
        IdProveedor INT NOT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_COM_CompraRetencion_TipoCambio DEFAULT (1),
        PorcentajeRetencion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_CompraRetencion_Porcentaje DEFAULT (0),
        Retencion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraRetencion_Importe DEFAULT (0),
        Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraRetencion_Saldo DEFAULT (0),
        ReferenciaDocumento NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_COM_CompraRetencion_Estado DEFAULT (N'PROVISIONADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_COM_CompraRetencion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.COM_CompraRetencion
        ADD CONSTRAINT FK_COM_CompraRetencion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.COM_CompraRetencion
        ADD CONSTRAINT FK_COM_CompraRetencion_COM_Compra
            FOREIGN KEY (IdCompra) REFERENCES dbo.COM_Compra (IdCompra);

    ALTER TABLE dbo.COM_CompraRetencion
        ADD CONSTRAINT FK_COM_CompraRetencion_ADM_Proveedor
            FOREIGN KEY (IdProveedor) REFERENCES dbo.ADM_Proveedor (IdProveedor);

    ALTER TABLE dbo.COM_CompraRetencion
        ADD CONSTRAINT FK_COM_CompraRetencion_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);
END;

MERGE dbo.CON_TipoImpuesto AS destino
USING
(
    SELECT
        CAST('R4TA' AS VARCHAR(10)) AS CodigoSunat,
        CAST(N'Renta de cuarta categoria por pagar' AS NVARCHAR(150)) AS NombreImpuesto
) AS fuente
    ON destino.CodigoSunat = fuente.CodigoSunat
WHEN MATCHED THEN
    UPDATE
    SET destino.NombreImpuesto = fuente.NombreImpuesto,
        destino.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, NombreImpuesto, CodigoCuenta, Estado)
    VALUES (fuente.CodigoSunat, fuente.NombreImpuesto, NULL, 1);

MERGE dbo.ADM_ParametroMaestro AS destino
USING
(
    SELECT
        CAST('NA' AS VARCHAR(50)) AS TipoParametro,
        CAST('PORCRETEN4TA' AS VARCHAR(50)) AS CodigoParametro,
        CAST('8' AS NVARCHAR(200)) AS ValorParametro,
        CAST(N'Porcentaje de retencion de renta de cuarta categoria para recibos por honorarios' AS NVARCHAR(500)) AS DescripcionParametro,
        CAST(NULL AS DATE) AS FecIni,
        CAST(NULL AS DATE) AS FecFin,
        CAST(286 AS INT) AS Orden
) AS fuente
    ON destino.TipoParametro = fuente.TipoParametro
   AND destino.CodigoParametro = fuente.CodigoParametro
WHEN MATCHED THEN
    UPDATE
    SET destino.ValorParametro = fuente.ValorParametro,
        destino.DescripcionParametro = fuente.DescripcionParametro,
        destino.Orden = fuente.Orden,
        destino.Activo = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (TipoParametro, CodigoParametro, ValorParametro, DescripcionParametro, FecIni, FecFin, Orden, Activo)
    VALUES (fuente.TipoParametro, fuente.CodigoParametro, fuente.ValorParametro, fuente.DescripcionParametro, fuente.FecIni, fuente.FecFin, fuente.Orden, 1);
