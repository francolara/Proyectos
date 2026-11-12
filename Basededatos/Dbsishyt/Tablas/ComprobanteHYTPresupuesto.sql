-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registro operativo de compras y comprobantes de venta.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ComprobanteHYTPresupuesto',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ComprobanteHYTPresupuesto
    (
        ComprobanteId BIGINT NOT NULL REFERENCES dbo.ComprobanteHYT(ComprobanteId),
        PresupuestoId BIGINT NOT NULL REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT PK_ComprobanteHYTPresupuesto PRIMARY KEY (ComprobanteId, PresupuestoId)
    );
END;
GO
