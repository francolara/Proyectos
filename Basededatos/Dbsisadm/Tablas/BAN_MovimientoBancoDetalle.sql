-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea el detalle contable de movimientos de caja y bancos con debe, haber y referencias documentarias por linea.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Amplia el detalle de Caja y Bancos para guardar la persona por linea, el origen del comprobante y el importe aplicado al saldo.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega importes por moneda al detalle para conservar total en soles y dolares por linea del movimiento bancario.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Incorpora TotalImporteS y TotalImporteD para conservar equivalencias por moneda en cada linea bancaria.
-- Firma: FRANCO LARA - 29/06/2026 | Vuelve obligatorio TipoCambioLinea en el detalle bancario y deja default en 1 para nuevas lineas.

IF OBJECT_ID(N'dbo.BAN_MovimientoBancoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BAN_MovimientoBancoDetalle
    (
        IdMovimientoBancoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_BAN_MovimientoBancoDetalle PRIMARY KEY,
        IdMovimientoBanco INT NOT NULL,
        Item SMALLINT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        IdPersona INT NULL,
        ModuloOperacionComprobante CHAR(3) NULL,
        IdRegistroComprobante INT NULL,
        ImporteAplicado DECIMAL(18, 2) NULL,
        GlosaDetalle NVARCHAR(300) NULL,
        CodigoCentroCosto VARCHAR(20) NULL,
        NumeroDocumento VARCHAR(20) NULL,
        TipoDocumento NVARCHAR(150) NULL,
        Serie VARCHAR(10) NULL,
        ReferenciaLinea NVARCHAR(100) NULL,
        TipoCambioLinea DECIMAL(18, 6) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TipoCambioLinea DEFAULT (1),
        Debe DECIMAL(18, 2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_Debe DEFAULT (0),
        Haber DECIMAL(18, 2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_Haber DEFAULT (0),
        TotalImporteS DECIMAL(18, 2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteS DEFAULT (0),
        TotalImporteD DECIMAL(18, 2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteD DEFAULT (0)
    );

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT FK_BAN_MovimientoBancoDetalle_BAN_MovimientoBanco
            FOREIGN KEY (IdMovimientoBanco) REFERENCES dbo.BAN_MovimientoBanco (IdMovimientoBanco);

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT FK_BAN_MovimientoBancoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT FK_BAN_MovimientoBancoDetalle_ADM_Persona
            FOREIGN KEY (IdPersona) REFERENCES dbo.ADM_Persona (IdPersona);

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT CK_BAN_MovimientoBancoDetalle_Debe
            CHECK (Debe >= 0);

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT CK_BAN_MovimientoBancoDetalle_Haber
            CHECK (Haber >= 0);

    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT UQ_BAN_MovimientoBancoDetalle_Item
            UNIQUE (IdMovimientoBanco, Item);
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'ModuloOperacionComprobante') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD ModuloOperacionComprobante CHAR(3) NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'IdRegistroComprobante') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD IdRegistroComprobante INT NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'ImporteAplicado') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD ImporteAplicado DECIMAL(18,2) NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'TotalImporteS') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteS DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'TotalImporteD') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TotalImporteD DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBancoDetalle', N'TipoCambioLinea') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD TipoCambioLinea DECIMAL(18,6) NOT NULL CONSTRAINT DF_BAN_MovimientoBancoDetalle_TipoCambioLinea DEFAULT (1);
END;

UPDATE d
SET TipoCambioLinea = ISNULL(NULLIF(d.TipoCambioLinea, 0), CASE WHEN m.TipoCambio > 0 THEN m.TipoCambio ELSE 1 END)
FROM dbo.BAN_MovimientoBancoDetalle AS d
INNER JOIN dbo.BAN_MovimientoBanco AS m
    ON m.IdMovimientoBanco = d.IdMovimientoBanco
WHERE d.TipoCambioLinea IS NULL
   OR d.TipoCambioLinea <= 0;

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.BAN_MovimientoBancoDetalle')
      AND name = N'TipoCambioLinea'
      AND is_nullable = 1
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ALTER COLUMN TipoCambioLinea DECIMAL(18,6) NOT NULL;
END;
