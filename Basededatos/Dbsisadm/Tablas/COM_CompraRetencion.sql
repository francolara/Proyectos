-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Documento pendiente de pago por retencion de renta de 4ta originado desde una compra de recibo por honorarios.
-- =============================================

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

    ALTER TABLE dbo.COM_CompraRetencion
        ADD CONSTRAINT CK_COM_CompraRetencion_Montos
            CHECK (PorcentajeRetencion >= 0 AND Retencion >= 0 AND Saldo >= 0);
END;
