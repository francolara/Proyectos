-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Detalle de cuentas y montos por asiento contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Agrega datos documentarios opcionales por linea para el registro manual.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Amplia TipoDocumento para guardar descripciones de comprobante en asientos automaticos de compras y ventas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega importes por moneda al detalle de asiento para conservar conversiones a soles y dolares por linea.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Incorpora TotalImporteS y TotalImporteD para conservar equivalencias por moneda en cada linea del asiento.
-- Firma: FRANCO LARA - 29/06/2026 | Vuelve obligatorio TipoCambioLinea en el detalle del asiento y deja default en 1 para nuevas lineas.
-- Firma: FRANCO LARA - 03/07/2026 | Agrega DH al detalle contable para guardar explicitamente el sentido Debe/Haber y alinea la restriccion de montos con esa marca.
-- Firma: FRANCO LARA - 06/07/2026 | Permite lineas analiticas de ajuste cambiario con Debe/Haber en cero cuando el saldo se conserva en TotalImporteS y/o TotalImporteD.

IF OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AsientoDetalle
    (
        IdAsientoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AsientoDetalle PRIMARY KEY,
        IdAsiento INT NOT NULL,
        Item SMALLINT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        GlosaDetalle NVARCHAR(300) NULL,
        CodigoCentroCosto NVARCHAR(50) NULL,
        TipoDocumento NVARCHAR(150) NULL,
        NumeroDocumento VARCHAR(20) NULL,
        Serie VARCHAR(10) NULL,
        TipoCambioLinea DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TipoCambioLinea DEFAULT (1),
        IdCliente INT NULL,
        IdProveedor INT NULL,
        DH CHAR(1) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_DH DEFAULT ('D'),
        Debe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_Debe DEFAULT (0),
        Haber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_Haber DEFAULT (0),
        TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteS DEFAULT (0),
        TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteD DEFAULT (0),
        ReferenciaLinea NVARCHAR(100) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT FK_CON_AsientoDetalle_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT FK_CON_AsientoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT FK_CON_AsientoDetalle_ADM_Cliente
            FOREIGN KEY (IdCliente) REFERENCES dbo.ADM_Cliente (IdCliente);

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT FK_CON_AsientoDetalle_ADM_Proveedor
            FOREIGN KEY (IdProveedor) REFERENCES dbo.ADM_Proveedor (IdProveedor);

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT CK_CON_AsientoDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT CK_CON_AsientoDetalle_Montos
            CHECK (
                Debe >= 0
                AND Haber >= 0
                AND (
                    (Debe > 0 AND Haber = 0)
                    OR (Debe = 0 AND Haber > 0)
                )
            );

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT UQ_CON_AsientoDetalle_IdAsiento_Item
            UNIQUE (IdAsiento, Item);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'TotalImporteS') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteS DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'TotalImporteD') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TotalImporteD DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'TipoCambioLinea') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TipoCambioLinea DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_TipoCambioLinea DEFAULT (1);
END;

IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'DH') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD DH CHAR(1) NULL;
END;

UPDATE d
SET DH = CASE
             WHEN d.Debe > 0 THEN 'D'
             WHEN d.Haber > 0 THEN 'H'
             ELSE ISNULL(d.DH, 'D')
         END
FROM dbo.CON_AsientoDetalle AS d
WHERE d.DH IS NULL
   OR d.DH NOT IN ('D', 'H');

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
      AND name = N'DH'
      AND is_nullable = 1
)
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ALTER COLUMN DH CHAR(1) NOT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
      AND name = N'DF_CON_AsientoDetalle_DH'
)
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT DF_CON_AsientoDetalle_DH DEFAULT ('D') FOR DH;
END;

UPDATE d
SET TipoCambioLinea = ISNULL(NULLIF(d.TipoCambioLinea, 0), CASE WHEN a.TipoCambio > 0 THEN a.TipoCambio ELSE 1 END)
FROM dbo.CON_AsientoDetalle AS d
INNER JOIN dbo.CON_Asiento AS a
    ON a.IdAsiento = d.IdAsiento
WHERE d.TipoCambioLinea IS NULL
   OR d.TipoCambioLinea <= 0;

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
      AND name = N'TipoCambioLinea'
      AND is_nullable = 1
)
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ALTER COLUMN TipoCambioLinea DECIMAL(18,6) NOT NULL;
END;

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
      AND name = N'CK_CON_AsientoDetalle_Montos'
)
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        DROP CONSTRAINT CK_CON_AsientoDetalle_Montos;
END;

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT CK_CON_AsientoDetalle_Montos
            CHECK (
                DH IN ('D', 'H')
                AND Debe >= 0
                AND Haber >= 0
                AND (
                    (DH = 'D' AND Debe > 0 AND Haber = 0)
                    OR (DH = 'H' AND Debe = 0 AND Haber > 0)
                    OR (
                        Debe = 0
                        AND Haber = 0
                        AND (
                            TotalImporteS > 0
                            OR TotalImporteD > 0
                        )
                    )
                )
            );
