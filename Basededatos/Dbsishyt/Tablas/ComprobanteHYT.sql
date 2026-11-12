-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registro operativo de compras y comprobantes de venta.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Conserva la tasa de detracción y la orden de compra vinculada en ventas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ComprobanteHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ComprobanteHYT
    (
        ComprobanteId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ComprobanteHYT PRIMARY KEY,
        OperacionId UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() CONSTRAINT UQ_ComprobanteHYT_Operacion UNIQUE,
        Huella VARBINARY(32) NOT NULL,
        EsVenta BIT NOT NULL,
        Tipo CHAR(2) NOT NULL CHECK (Tipo IN ('01','03','07')),
        Serie NVARCHAR(30) NULL,
        Numero NVARCHAR(50) NULL,
        Fecha DATE NOT NULL,
        Moneda NVARCHAR(3) NOT NULL REFERENCES dbo.Moneda(Codigo),
        RucHYT VARCHAR(11) NOT NULL,
        DocumentoTercero NVARCHAR(30) NOT NULL,
        NombreTercero NVARCHAR(180) NOT NULL,
        ProveedorId BIGINT NULL REFERENCES dbo.ProveedorHYT(ProveedorId),
        ClienteId BIGINT NULL REFERENCES dbo.Cliente(ClienteId),
        OrdenCompraId BIGINT NULL CONSTRAINT FK_ComprobanteHYT_OrdenCompra REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        FormaPago NVARCHAR(30) NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CHECK (Importe >= 0),
        Detalle NVARCHAR(MAX) NOT NULL CHECK (ISJSON(Detalle)=1),
        ImporteDetraccion DECIMAL(18,2) NOT NULL CHECK (ImporteDetraccion >= 0),
        DetraccionDocumento DECIMAL(18,2) NOT NULL CHECK (DetraccionDocumento >= 0),
        MonedaDetraccion NVARCHAR(3) NOT NULL REFERENCES dbo.Moneda(Codigo),
        CodigoDetraccion NVARCHAR(30) NULL,
        PorcentajeDetraccion DECIMAL(7,4) NULL CONSTRAINT CK_ComprobanteHYT_PorcentajeDetraccion CHECK (PorcentajeDetraccion BETWEEN 0 AND 100),
        ReferenciaNota NVARCHAR(100) NULL,
        MotivoNota NVARCHAR(300) NULL,
        XmlOriginal NVARCHAR(MAX) NULL,
        EstadoSunat NVARCHAR(20) NOT NULL DEFAULT N'PENDIENTE',
        ProveedorElectronico NVARCHAR(60) NULL,
        IdentificadorElectronico NVARCHAR(150) NULL,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_ComprobanteHYT_Tercero CHECK ((EsVenta=0 AND ProveedorId IS NOT NULL AND ClienteId IS NULL) OR (EsVenta=1 AND ClienteId IS NOT NULL AND ProveedorId IS NULL))
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.ComprobanteHYT') AND name=N'UX_ComprobanteHYT_Documento')
    CREATE UNIQUE INDEX UX_ComprobanteHYT_Documento ON dbo.ComprobanteHYT(EsVenta,RucHYT,DocumentoTercero,Tipo,Serie,Numero)
        WHERE Serie IS NOT NULL AND Numero IS NOT NULL;
GO
