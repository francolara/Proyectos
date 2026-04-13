USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 09/04/2026 | Ajuste a CREATE OR ALTER y salida de codigo de documento para UI de comprobantes.
-- Firma: Codex - 11/04/2026 | Incluye datos de referencia/tipo de nota y codigos 07/08 para NC/ND.
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
            c.TipoComprobante,
            c.Serie,
            c.Numero,
            c.FechaEmision,
            c.TipoMoneda,
            c.SubTotal,
            c.Igv,
            c.Total,
            c.Estado,
            CASE
                WHEN c.TipoComprobante = 2 THEN N'01'
                WHEN c.TipoComprobante = 1 THEN N'03'
                WHEN c.TipoComprobante = 3 THEN N'RI'
                WHEN c.TipoComprobante = 4 THEN N'07'
                WHEN c.TipoComprobante = 5 THEN N'08'
                ELSE N'03'
            END AS CodigoDocumentoComprobante,
            c.ComprobanteReferenciaId,
            c.TipoNota,
            c.TipoNotaCodigoSunat
        FROM dbo.ComprobantesElectronicos c
        WHERE c.NegocioId = @NegocioId
          AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
