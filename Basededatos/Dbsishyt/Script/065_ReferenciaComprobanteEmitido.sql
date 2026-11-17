-- Firma: FRANCO LARA | Create date: 17/11/2026 | Distingue referencias manuales y emitidas y mantiene los requisitos para las manuales.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ReferenciaComprobanteExterno',N'ComprobanteId') IS NULL
    ALTER TABLE dbo.ReferenciaComprobanteExterno ADD ComprobanteId BIGINT NULL;
GO
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.ReferenciaComprobanteExterno')
        AND name=N'ConfirmacionId' AND is_nullable=0)
    ALTER TABLE dbo.ReferenciaComprobanteExterno ALTER COLUMN ConfirmacionId BIGINT NULL;
GO
IF OBJECT_ID(N'dbo.CK_ReferenciaComprobante_Origen',N'C') IS NULL
    ALTER TABLE dbo.ReferenciaComprobanteExterno ADD CONSTRAINT CK_ReferenciaComprobante_Origen
        CHECK ((ComprobanteId IS NULL AND ConfirmacionId IS NOT NULL)
            OR (ComprobanteId IS NOT NULL AND ConfirmacionId IS NULL AND Vigente=1));
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
        WHERE referencia.Vigente = 1 AND referencia.ComprobanteId IS NULL AND
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
