USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Comprobantes_ObtenerPorId]    Script Date: 5/05/2026 14:02:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 09/04/2026 | Ajuste a CREATE OR ALTER y salida de codigo de documento para UI de comprobantes.
-- Firma: Codex - 11/04/2026 | Incluye datos de referencia/tipo de nota y codigos 07/08 para NC/ND.
ALTER   PROCEDURE [dbo].[Sp_Comprobantes_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.ReservaId,
            c.TipoComprobante,
            c.Serie,
            c.Numero,
            c.FechaEmision,
            c.TipoMoneda,
            c.SubTotal,
            c.Igv,
            c.Total,
            c.Estado,
            d.CodigoSunat  AS CodigoDocumentoComprobante,
            c.ComprobanteReferenciaId,
            c.TipoNota,
            c.TipoNotaCodigoSunat,

            CASE
                WHEN  d.CodigoSunat = '01' THEN 1 -- Factura
                WHEN  d.CodigoSunat = '03' THEN 2 -- Boleta
                WHEN  d.CodigoSunat = 'RI' THEN 0
                WHEN  d.CodigoSunat = '07' THEN 3 -- Nota de Credito
                WHEN  d.CodigoSunat = '08' THEN 4 -- Nota de Debito
            END AS CodigoDocumentoComprobantenb,
            CASE WHEN ltrim(rtrim(isnull(e.CodigoUbigeo,'')))  = '' THEN F.CodigoUbigeo ELSE ltrim(rtrim(isnull(e.CodigoUbigeo,''))) END AS ClienteCodigoUbigeo,
            CASE WHEN ISNULL(e.TipoDocumento,0) = 0 THEN '-' ELSE ISNULL(e.TipoDocumento,0) END AS ClienteTipoDocumento,
            CASE WHEN ISNULL(e.TipoDocumento,0) = 0 THEN '-' ELSE e.NumeroDocumento END AS ClienteNumeroDocumento,
            CASE WHEN M.Codigo = 'PEN' THEN 1 
                 WHEN M.Codigo = 'USD' THEN 2 END MonedaNubefact 
            
        FROM dbo.ComprobantesElectronicos c
        inner join NegociosTiposDocumentoComprobante d
        ON c.TipoComprobante = d.Id AND d.NegocioId = @NegocioId
        INNER JOIN clientes e
        ON c.ClienteId = e.Id
        AND c.NegocioId = e.NegocioId
        INNER JOIN Negocios f
        on c.NegocioId = f.Id
        INNER JOIN Monedas M
        on m.NegocioId = c.NegocioId
        and m.Id = c.TipoMoneda
        WHERE c.NegocioId = @NegocioId
        AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
