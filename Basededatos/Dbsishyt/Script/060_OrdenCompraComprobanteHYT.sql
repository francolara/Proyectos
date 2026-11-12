-- Firma: FRANCO LARA | Create date: 11/11/2026 | Vincula una orden de compra del cliente al comprobante de venta.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.ComprobanteHYT',N'OrdenCompraId') IS NULL
    ALTER TABLE dbo.ComprobanteHYT ADD OrdenCompraId BIGINT NULL
        CONSTRAINT FK_ComprobanteHYT_OrdenCompra REFERENCES dbo.OrdenCompraCliente(OrdenCompraId);
GO
