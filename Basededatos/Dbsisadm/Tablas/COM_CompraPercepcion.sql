-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Documento pendiente de pago por percepcion originado desde una compra.
-- =============================================

IF OBJECT_ID(N'dbo.COM_CompraPercepcion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.COM_CompraPercepcion
    (
        IdCompraPercepcion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_COM_CompraPercepcion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdCompra INT NOT NULL,
        IdProveedor INT NOT NULL,
        IdTipoPercepcion INT NOT NULL,
        IdAsiento INT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_TipoCambio DEFAULT (1),
        CodigoPercepcion VARCHAR(2) NOT NULL,
        DescripcionPercepcion NVARCHAR(200) NOT NULL,
        PorcentajePercepcion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_Porcentaje DEFAULT (0),
        BasePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_Base DEFAULT (0),
        ImportePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_Importe DEFAULT (0),
        Saldo DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_Saldo DEFAULT (0),
        ReferenciaDocumento NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_Estado DEFAULT (N'PROVISIONADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_COM_CompraPercepcion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_COM_Compra
            FOREIGN KEY (IdCompra) REFERENCES dbo.COM_Compra (IdCompra);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_ADM_Proveedor
            FOREIGN KEY (IdProveedor) REFERENCES dbo.ADM_Proveedor (IdProveedor);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_ADM_TipoPercepcion
            FOREIGN KEY (IdTipoPercepcion) REFERENCES dbo.ADM_TipoPercepcion (IdTipoPercepcion);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT FK_COM_CompraPercepcion_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.COM_CompraPercepcion
        ADD CONSTRAINT CK_COM_CompraPercepcion_Montos
            CHECK (PorcentajePercepcion >= 0 AND BasePercepcion >= 0 AND ImportePercepcion >= 0 AND Saldo >= 0);
END;
