
GO
/****** Object:  StoredProcedure [dbo].[Sp_Comprobantes_ObtenerPorId]    Script Date: 5/05/2026 14:02:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 09/04/2026 | Ajuste a CREATE OR ALTER y salida de codigo de documento para UI de comprobantes.
-- Firma: Codex - 11/04/2026 | Incluye datos de referencia/tipo de nota y codigos 07/08 para NC/ND.
-- Firma: FRANCO LARA - 17/09/2026 | Incluye trazabilidad del comprobante para su visualizacion durante la edicion.
-- Firma: FRANCO LARA - 06/10/2026 | Expone codigos canonicos y reemplaza los ordinales heredados por valores reservados nulos.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Comprobantes_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.ReservaId,
            CAST(NULL AS INT) AS CompatibilidadTipoComprobanteReservada,
            c.Serie,
            c.Numero,
            c.FechaEmision,
            CAST(NULL AS INT) AS CompatibilidadTipoMonedaReservada,
            c.SubTotal,
            c.Igv,
            c.Total,
            c.Estado,
            c.CodigoTipoComprobante AS CodigoDocumentoComprobante,
            c.ComprobanteReferenciaId,
            c.TipoNota,
            c.TipoNotaCodigoSunat,

            CASE
                WHEN c.CodigoTipoComprobante = '01' THEN 1
                WHEN c.CodigoTipoComprobante = '03' THEN 2
                WHEN c.CodigoTipoComprobante = 'RI' THEN 0
                WHEN c.CodigoTipoComprobante = '07' THEN 3
                WHEN c.CodigoTipoComprobante = '08' THEN 4
            END AS CodigoDocumentoComprobantenb,
            CASE WHEN ltrim(rtrim(isnull(e.CodigoUbigeo,'')))  = '' THEN F.CodigoUbigeo ELSE ltrim(rtrim(isnull(e.CodigoUbigeo,''))) END AS ClienteCodigoUbigeo,
            CASE WHEN ISNULL(e.TipoDocumento,0) = 0 THEN '-' ELSE ISNULL(e.TipoDocumento,0) END AS ClienteTipoDocumento,
            CASE WHEN ISNULL(e.TipoDocumento,0) = 0 THEN '-' ELSE e.NumeroDocumento END AS ClienteNumeroDocumento,
            CASE WHEN c.CodigoMoneda = 'PEN' THEN 1
                 WHEN c.CodigoMoneda = 'USD' THEN 2 END MonedaNubefact,
            c.UsuarioCreacion,
            CAST(c.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaRegistro,
            c.UsuarioActualizacion,
            CAST(c.FechaActualizacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaActualizacion,
            c.CodigoMoneda

        FROM dbo.ComprobantesElectronicos c
        INNER JOIN clientes e
        ON c.ClienteId = e.Id
        AND c.NegocioId = e.NegocioId
        INNER JOIN Negocios f
        on c.NegocioId = f.Id
        WHERE c.NegocioId = @NegocioId
        AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
