-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Documento pendiente de detraccion generado desde la provision de compras para controlar su saldo y asiento independiente.
-- =============================================

IF OBJECT_ID(N'dbo.COM_CompraDetraccion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.COM_CompraDetraccion
    (
        IdCompraDetraccion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_COM_CompraDetraccion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdCompra INT NOT NULL,
        IdProveedor INT NOT NULL,
        IdDetraccionSunat INT NOT NULL,
        IdAsiento INT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_COM_CompraDetraccion_TipoCambio DEFAULT (1),
        CodigoDetraccionSunat VARCHAR(3) NOT NULL,
        DescripcionDetraccion NVARCHAR(250) NOT NULL,
        PorcentajeDetraccion DECIMAL(7,4) NOT NULL,
        ImporteDetraccion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraDetraccion_Importe DEFAULT (0),
        Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraDetraccion_Saldo DEFAULT (0),
        ReferenciaDocumento NVARCHAR(50) NOT NULL,
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_COM_CompraDetraccion_Estado DEFAULT (N'PROVISIONADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_COM_CompraDetraccion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_COM_Compra
            FOREIGN KEY (IdCompra) REFERENCES dbo.COM_Compra (IdCompra);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_ADM_Proveedor
            FOREIGN KEY (IdProveedor) REFERENCES dbo.ADM_Proveedor (IdProveedor);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_ADM_DetraccionSunat
            FOREIGN KEY (IdDetraccionSunat) REFERENCES dbo.ADM_DetraccionSunat (IdDetraccionSunat);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT FK_COM_CompraDetraccion_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT UQ_COM_CompraDetraccion_Compra
            UNIQUE (IdCompra);

    ALTER TABLE dbo.COM_CompraDetraccion
        ADD CONSTRAINT CK_COM_CompraDetraccion_Montos
            CHECK (ImporteDetraccion >= 0 AND Saldo >= 0 AND Saldo <= ImporteDetraccion AND PorcentajeDetraccion >= 0 AND PorcentajeDetraccion <= 100);
END;
