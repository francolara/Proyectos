-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Agrega totales exonerado/inafecto e ICBPER interno en compras, cuenta contable y tipo de afectacion IGV en detalle de compra.
-- =============================================

IF OBJECT_ID(N'dbo.CON_TipoAfectacionIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoAfectacionIGV
    (
        IdTipoAfectacionIGV INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoAfectacionIGV PRIMARY KEY,
        CodigoSunat VARCHAR(10) NOT NULL,
        NombreAfectacion NVARCHAR(120) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_TipoAfectacionIGV_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_TipoAfectacionIGV
        ADD CONSTRAINT UQ_CON_TipoAfectacionIGV_CodigoSunat
            UNIQUE (CodigoSunat);
END;

MERGE dbo.CON_TipoAfectacionIGV AS destino
USING
(
    VALUES
        ('10', N'Gravado - Operacion Onerosa', 1),
        ('20', N'Exonerado - Operacion Onerosa', 1),
        ('30', N'Inafecto - Operacion Onerosa', 1),
        ('40', N'Exportacion', 1),
        ('21', N'Exonerado - Transferencia Gratuita', 1),
        ('31', N'Inafecto - Transferencia Gratuita', 1)
) AS fuente (CodigoSunat, NombreAfectacion, Estado)
    ON destino.CodigoSunat = fuente.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        NombreAfectacion = fuente.NombreAfectacion,
        Estado = fuente.Estado
WHEN NOT MATCHED BY TARGET THEN
    INSERT
    (
        CodigoSunat,
        NombreAfectacion,
        Estado
    )
    VALUES
    (
        fuente.CodigoSunat,
        fuente.NombreAfectacion,
        fuente.Estado
    );

IF COL_LENGTH(N'dbo.COM_Compra', N'Icbper') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD Icbper DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_Icbper DEFAULT (0);
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

IF COL_LENGTH(N'dbo.COM_CompraDetalle', N'IdPlanCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.COM_CompraDetalle
        ADD IdPlanCuenta INT NULL;
END;

IF COL_LENGTH(N'dbo.COM_CompraDetalle', N'IdTipoAfectacionIGV') IS NULL
BEGIN
    ALTER TABLE dbo.COM_CompraDetalle
        ADD IdTipoAfectacionIGV INT NULL;
END;

IF OBJECT_ID(N'dbo.FK_COM_CompraDetalle_CON_PlanCuenta', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT FK_COM_CompraDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.FK_COM_CompraDetalle_CON_TipoAfectacionIGV', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT FK_COM_CompraDetalle_CON_TipoAfectacionIGV
            FOREIGN KEY (IdTipoAfectacionIGV) REFERENCES dbo.CON_TipoAfectacionIGV (IdTipoAfectacionIGV);
END;
