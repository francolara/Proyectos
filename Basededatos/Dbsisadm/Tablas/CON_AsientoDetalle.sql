-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Detalle de cuentas y montos por asiento contable.
-- =============================================

IF OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AsientoDetalle
    (
        IdAsientoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AsientoDetalle PRIMARY KEY,
        IdAsiento INT NOT NULL,
        Item SMALLINT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        GlosaDetalle NVARCHAR(300) NULL,
        IdCliente INT NULL,
        IdProveedor INT NULL,
        Debe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_Debe DEFAULT (0),
        Haber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AsientoDetalle_Haber DEFAULT (0),
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
