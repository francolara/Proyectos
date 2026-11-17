-- Firma: FRANCO LARA | Create date: 17/11/2026 | Sincroniza altas, cambios, bajas, notas de crédito y vínculos de facturas.
USE [SistemaWebHYT];
GO
CREATE OR ALTER TRIGGER dbo.tr_ComprobanteHYT_FacturadoOC
ON dbo.ComprobanteHYT AFTER INSERT,UPDATE,DELETE
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Ordenes NVARCHAR(MAX), @Motivo NVARCHAR(1000);
    WITH cambios AS (
        SELECT i.ComprobanteId,i.OrdenCompraId,i.Tipo,i.ReferenciaNota,i.ClienteId,i.RucHYT FROM inserted AS i WHERE i.EsVenta=1
        UNION
        SELECT d.ComprobanteId,d.OrdenCompraId,d.Tipo,d.ReferenciaNota,d.ClienteId,d.RucHYT FROM deleted AS d WHERE d.EsVenta=1
    ), facturas AS (
        SELECT cambio.ComprobanteId,cambio.OrdenCompraId FROM cambios AS cambio WHERE cambio.Tipo IN ('01','03')
        UNION
        SELECT documento.ComprobanteId,documento.OrdenCompraId
        FROM cambios AS nota INNER JOIN dbo.ComprobanteHYT AS documento
            ON documento.EsVenta=1 AND documento.Tipo IN ('01','03') AND documento.ClienteId=nota.ClienteId
            AND documento.RucHYT=nota.RucHYT AND CONCAT(documento.Serie,N'-',documento.Numero)=LTRIM(RTRIM(nota.ReferenciaNota))
        WHERE nota.Tipo='07'
    ), afectadas AS (
        SELECT factura.OrdenCompraId FROM facturas AS factura
        UNION
        SELECT orden.OrdenCompraId FROM facturas AS factura
        INNER JOIN dbo.ComprobanteHYTPresupuesto AS vinculo ON vinculo.ComprobanteId=factura.ComprobanteId
        INNER JOIN dbo.OrdenCompraCliente AS orden ON orden.PresupuestoId=vinculo.PresupuestoId
    )
    SELECT @Ordenes=CONCAT(N'[',STRING_AGG(CONVERT(NVARCHAR(MAX),afectada.OrdenCompraId),N','),N']')
    FROM afectadas AS afectada;
    IF @Ordenes=N'[]' RETURN;
    SET @Motivo=CASE WHEN NOT EXISTS(SELECT 1 FROM inserted) THEN N'Comprobante eliminado.'
        WHEN EXISTS(SELECT 1 FROM inserted AS i WHERE i.Tipo='07') THEN N'Comprobante anulado en su totalidad por nota de crédito.'
        ELSE N'Comprobante sin vigencia o vínculo modificado.' END;
    EXEC dbo.spu_Facturado_SincronizarEmision @Ordenes,@Motivo;
END;
GO
CREATE OR ALTER TRIGGER dbo.tr_ComprobanteHYTPresupuesto_FacturadoOC
ON dbo.ComprobanteHYTPresupuesto AFTER INSERT,UPDATE,DELETE
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Ordenes NVARCHAR(MAX);
    WITH afectados AS (
        SELECT i.PresupuestoId FROM inserted AS i
        UNION
        SELECT d.PresupuestoId FROM deleted AS d
    )
    SELECT @Ordenes=CONCAT(N'[',STRING_AGG(CONVERT(NVARCHAR(MAX),orden.OrdenCompraId),N','),N']')
    FROM afectados AS afectado
    INNER JOIN dbo.OrdenCompraCliente AS orden ON orden.PresupuestoId=afectado.PresupuestoId;
    IF @Ordenes<>N'[]'
        EXEC dbo.spu_Facturado_SincronizarEmision @Ordenes,N'Comprobante desvinculado del presupuesto.';
END;
GO
CREATE OR ALTER TRIGGER dbo.tr_OrdenCompraCliente_FacturadoEmision
ON dbo.OrdenCompraCliente AFTER INSERT,UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Ordenes NVARCHAR(MAX);
    SELECT @Ordenes=CONCAT(N'[',STRING_AGG(CONVERT(NVARCHAR(MAX),i.OrdenCompraId),N','),N']')
    FROM inserted AS i LEFT JOIN deleted AS d ON d.OrdenCompraId=i.OrdenCompraId
    WHERE d.OrdenCompraId IS NULL OR (i.Estado<>d.Estado AND (i.Estado=N'CANCELADA' OR d.Estado=N'CANCELADA'));
    IF @Ordenes<>N'[]'
        EXEC dbo.spu_Facturado_SincronizarEmision @Ordenes,N'Orden de compra cancelada.';
END;
GO
EXEC dbo.spu_Facturado_SincronizarEmision;
GO
