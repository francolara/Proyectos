-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Detalle referencial de conceptos de la compra provisionada.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Agrega cuenta contable y tipo de afectacion IGV por linea de compra.
-- =============================================

IF OBJECT_ID(N'dbo.COM_CompraDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.COM_CompraDetalle
    (
        IdCompraDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_COM_CompraDetalle PRIMARY KEY,
        IdCompra INT NOT NULL,
        Item SMALLINT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        IdTipoAfectacionIGV INT NOT NULL,
        Descripcion NVARCHAR(250) NOT NULL,
        Cantidad DECIMAL(18,4) NOT NULL CONSTRAINT DF_COM_CompraDetalle_Cantidad DEFAULT (1),
        ValorUnitario DECIMAL(18,6) NOT NULL CONSTRAINT DF_COM_CompraDetalle_ValorUnitario DEFAULT (0),
        ImporteBruto DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraDetalle_ImporteBruto DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_COM_CompraDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT FK_COM_CompraDetalle_COM_Compra
            FOREIGN KEY (IdCompra) REFERENCES dbo.COM_Compra (IdCompra);

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT FK_COM_CompraDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT FK_COM_CompraDetalle_CON_TipoAfectacionIGV
            FOREIGN KEY (IdTipoAfectacionIGV) REFERENCES dbo.CON_TipoAfectacionIGV (IdTipoAfectacionIGV);

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT CK_COM_CompraDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT CK_COM_CompraDetalle_Montos
            CHECK (Cantidad > 0 AND ValorUnitario >= 0 AND ImporteBruto >= 0);

    ALTER TABLE dbo.COM_CompraDetalle
        ADD CONSTRAINT UQ_COM_CompraDetalle_Item
            UNIQUE (IdCompra, Item);
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
