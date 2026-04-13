/*
Firma: Codex - 13/04/2026
Descripcion: Ajusta indice por reserva en comprobantes para permitir reemision cuando existe NC activa sin anular comprobante inicial.
*/
USE [DbSportCenter]
GO

SET NOCOUNT ON;

IF EXISTS
(
    SELECT 1
    FROM sys.indexes i
    WHERE i.object_id = OBJECT_ID(N'dbo.ComprobantesElectronicos')
      AND i.name = N'IX_ComprobantesElectronicos_ReservaId'
)
BEGIN
    DROP INDEX [IX_ComprobantesElectronicos_ReservaId] ON [dbo].[ComprobantesElectronicos];
END;

/*
Regla:
- Se mantiene indice filtrado para comprobantes principales activos.
- Se elimina restriccion UNIQUE porque la regla de reemision por NC activa
  se controla en SP (Sp_Comprobantes_Crear) y no puede expresarse en filtro de indice.
*/
CREATE NONCLUSTERED INDEX [IX_ComprobantesElectronicos_ReservaId]
ON [dbo].[ComprobantesElectronicos]([NegocioId] ASC, [ReservaId] ASC)
WHERE [Estado] <> 5
  AND [ComprobanteReferenciaId] IS NULL;
GO
