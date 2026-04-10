USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Combo de series configuradas por negocio y tipo de documento.
-- Firma:         Codex - 09/04/2026 | Restringe combo a documentos activos en Maestros y supermaestro habilitado.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_SeriesDocumentoComprobante
    @NegocioId INT,
    @CodigoSunat NVARCHAR(4)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @CodigoSunat = UPPER(LTRIM(RTRIM(@CodigoSunat)));

        SELECT
            CONVERT(NVARCHAR(20), ns.Id) AS Value,
            ns.Serie AS Text
        FROM dbo.NegociosSeriesDocumentoComprobante ns
        INNER JOIN dbo.NegociosTiposDocumentoComprobante ntd
            ON ntd.NegocioId = ns.NegocioId
           AND ntd.CodigoSunat = ns.CodigoSunat
           AND ntd.Activo = 1
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t
            ON t.CodigoSunat = ns.CodigoSunat
           AND t.Activo = 1
           AND t.Habilitado = 1
        WHERE ns.NegocioId = @NegocioId
          AND ns.CodigoSunat = @CodigoSunat
          AND ns.Activo = 1
        ORDER BY ns.Serie;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
