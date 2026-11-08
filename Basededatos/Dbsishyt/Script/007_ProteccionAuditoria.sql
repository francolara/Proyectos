-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Protege auditoria, vigencias IGV y requisitos del marcado Facturado.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Guía registrada sin estado de aceptación para el marcado Facturado. OC con trabajo completado acredita término sin ejecución separada; confirma vínculo y vigencia de cliente.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GO
CREATE OR ALTER TRIGGER dbo.tr_EventoAuditoria_Inmutable
ON dbo.EventoAuditoria
INSTEAD OF UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;
    THROW 51009, N'El historial de auditoria es inmutable.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_TasaIGV_SinSolapamiento
ON dbo.TasaIGV
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM dbo.TasaIGV AS primera WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.TasaIGV AS segunda WITH (UPDLOCK, HOLDLOCK)
            ON primera.FechaDesde < segunda.FechaDesde
            AND (primera.FechaHasta IS NULL OR segunda.FechaDesde < primera.FechaHasta)
    )
        THROW 51010, N'Las vigencias de IGV no pueden solaparse.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_ReferenciaComprobante_Requisitos
ON dbo.ReferenciaComprobanteExterno
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS referencia
        INNER JOIN dbo.OrdenCompraCliente AS oc WITH (UPDLOCK, HOLDLOCK) ON oc.OrdenCompraId = referencia.OrdenCompraId
        INNER JOIN dbo.ConfirmacionFacturacion AS confirmacion WITH (UPDLOCK, HOLDLOCK)
            ON confirmacion.ConfirmacionId = referencia.ConfirmacionId
        WHERE referencia.Vigente = 1 AND
        (
            confirmacion.Vigente = 0 OR oc.Estado = N'CANCELADA'
            OR NOT EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS guia WITH (UPDLOCK, HOLDLOCK)
                WHERE guia.PresupuestoId = oc.PresupuestoId)
            OR confirmacion.OrdenCompraId <> oc.OrdenCompraId
            OR (oc.Estado NOT IN (N'TRABAJO_COMPLETADO', N'CERRADA_OPERATIVAMENTE') AND
                (NOT EXISTS (SELECT 1 FROM dbo.Ejecucion AS ejecucion WITH (UPDLOCK, HOLDLOCK)
                    WHERE ejecucion.PresupuestoId = oc.PresupuestoId AND ejecucion.Estado = N'TERMINADA')
                OR EXISTS (SELECT 1 FROM dbo.Ejecucion AS pendiente WITH (UPDLOCK, HOLDLOCK)
                    WHERE pendiente.PresupuestoId = oc.PresupuestoId AND pendiente.Estado NOT IN (N'TERMINADA', N'CANCELADA'))))
        )
    )
        THROW 51013, N'El marcado Facturado requiere OC valida, trabajo terminado, guia registrada y confirmacion vigente.', 1;
END;
GO
CREATE OR ALTER TRIGGER dbo.tr_Confirmacion_ProtegerFacturado
ON dbo.ConfirmacionFacturacion
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS confirmacion
        INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia WITH (UPDLOCK, HOLDLOCK)
            ON referencia.ConfirmacionId = confirmacion.ConfirmacionId AND referencia.Vigente = 1
        WHERE confirmacion.Vigente = 0
    )
        THROW 51014, N'Retire primero el marcado Facturado con motivo y auditoria antes de revocar su confirmacion.', 1;
END;
GO
