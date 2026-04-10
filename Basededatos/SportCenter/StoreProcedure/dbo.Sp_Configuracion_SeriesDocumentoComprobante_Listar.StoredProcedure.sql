USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Lista series configuradas por negocio para documentos de comprobante.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Configuracion_SeriesDocumentoComprobante_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            ns.Id,
            ns.CodigoSunat,
            t.Nombre,
            t.Tributario,
            ns.Serie,
            ns.Activo
        FROM dbo.NegociosSeriesDocumentoComprobante ns
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ns.CodigoSunat
        WHERE ns.NegocioId = @NegocioId
        ORDER BY t.Orden, ns.Serie;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
