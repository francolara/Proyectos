-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Detalle referencial de conceptos de la venta registrada.
-- =============================================

IF OBJECT_ID(N'dbo.VEN_VentaDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VEN_VentaDetalle
    (
        IdVentaDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_VEN_VentaDetalle PRIMARY KEY,
        IdVenta INT NOT NULL,
        Item SMALLINT NOT NULL,
        Descripcion NVARCHAR(250) NOT NULL,
        Cantidad DECIMAL(18,4) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_Cantidad DEFAULT (1),
        ValorUnitario DECIMAL(18,6) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_ValorUnitario DEFAULT (0),
        ImporteBruto DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_ImporteBruto DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT FK_VEN_VentaDetalle_VEN_Venta
            FOREIGN KEY (IdVenta) REFERENCES dbo.VEN_Venta (IdVenta);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT CK_VEN_VentaDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT CK_VEN_VentaDetalle_Montos
            CHECK (Cantidad > 0 AND ValorUnitario >= 0 AND ImporteBruto >= 0);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT UQ_VEN_VentaDetalle_Item
            UNIQUE (IdVenta, Item);
END;
