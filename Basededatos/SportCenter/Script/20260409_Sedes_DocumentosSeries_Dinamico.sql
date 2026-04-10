USE [DbSportCenter]
GO
-- Firma: Codex - 09/04/2026 | Sedes: Documentos y series dinamicos segun Maestros (activo) y series activas configuradas en Negocio.

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_SeriesDocumentoComprobante_Listar
    @NegocioId INT,
    @SedeId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Sedes WHERE Id = @SedeId AND NegocioId = @NegocioId)
            RAISERROR('La sede no pertenece al negocio.', 16, 1);

        ;WITH DocsConSeries AS
        (
            SELECT
                ntd.CodigoSunat
            FROM dbo.NegociosTiposDocumentoComprobante ntd
            INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
            WHERE ntd.NegocioId = @NegocioId
              AND ntd.Activo = 1
              AND t.Activo = 1
              AND t.Habilitado = 1
              AND EXISTS
              (
                  SELECT 1
                  FROM dbo.NegociosSeriesDocumentoComprobante nsx
                  WHERE nsx.NegocioId = @NegocioId
                    AND nsx.CodigoSunat = ntd.CodigoSunat
                    AND nsx.Activo = 1
              )
            GROUP BY ntd.CodigoSunat
        )
        SELECT
            t.CodigoSunat,
            t.Nombre,
            t.Tributario,
            ss.NegocioSerieId,
            ns.Serie
        FROM DocsConSeries d
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = d.CodigoSunat
        LEFT JOIN dbo.SedesSeriesDocumentoComprobante ss
            ON ss.SedeId = @SedeId
           AND ss.CodigoSunat = d.CodigoSunat
           AND ss.Activo = 1
        LEFT JOIN dbo.NegociosSeriesDocumentoComprobante ns
            ON ns.Id = ss.NegocioSerieId
           AND ns.NegocioId = @NegocioId
           AND ns.CodigoSunat = d.CodigoSunat
           AND ns.Activo = 1
        ORDER BY t.Orden, t.CodigoSunat;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

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
