-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Agrega total exonerado, total inafecto e ICBPER interno en ventas, con cuenta contable y tipo de afectacion IGV por detalle.
-- =============================================

IF COL_LENGTH(N'dbo.VEN_Venta', N'TotalExonerado') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_Venta
        ADD TotalExonerado DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_TotalExonerado DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.VEN_Venta', N'TotalInafecto') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_Venta
        ADD TotalInafecto DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_TotalInafecto DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.VEN_Venta', N'Icbper') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_Venta
        ADD Icbper DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Icbper DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.VEN_VentaDetalle', N'IdPlanCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_VentaDetalle
        ADD IdPlanCuenta INT NULL;
END;

IF COL_LENGTH(N'dbo.VEN_VentaDetalle', N'IdTipoAfectacionIGV') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_VentaDetalle
        ADD IdTipoAfectacionIGV INT NULL;
END;

IF OBJECT_ID(N'dbo.FK_VEN_VentaDetalle_CON_PlanCuenta', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT FK_VEN_VentaDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.FK_VEN_VentaDetalle_CON_TipoAfectacionIGV', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT FK_VEN_VentaDetalle_CON_TipoAfectacionIGV
            FOREIGN KEY (IdTipoAfectacionIGV) REFERENCES dbo.CON_TipoAfectacionIGV (IdTipoAfectacionIGV);
END;
