-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registro operativo de compras y comprobantes de venta.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ComprobanteHYTCosto',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ComprobanteHYTCosto
    (
        ComprobanteId BIGINT NOT NULL REFERENCES dbo.ComprobanteHYT(ComprobanteId),
        CostoProveedorId BIGINT NOT NULL CONSTRAINT PK_ComprobanteHYTCosto PRIMARY KEY REFERENCES dbo.CostoProveedorPresupuesto(CostoProveedorId)
    );
END;
GO
