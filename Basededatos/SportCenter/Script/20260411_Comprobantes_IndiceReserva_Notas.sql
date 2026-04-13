/*
Firma: Codex - 11/04/2026
Descripcion: Ajusta indice de unicidad por negocio+reserva para comprobante principal, permitiendo documentos derivados (NC/ND).
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
- Un comprobante principal activo por negocio+reserva (Estado <> 5 y ComprobanteReferenciaId IS NULL).
- Los comprobantes derivados (NC/ND u otros futuros) tienen ComprobanteReferenciaId IS NOT NULL y no participan del indice unico.
- No depende de valores fijos de TipoComprobante.
*/
CREATE UNIQUE NONCLUSTERED INDEX [IX_ComprobantesElectronicos_ReservaId]
ON [dbo].[ComprobantesElectronicos]([NegocioId] ASC, [ReservaId] ASC)
WHERE [Estado] <> 5
  AND [ComprobanteReferenciaId] IS NULL;
GO
