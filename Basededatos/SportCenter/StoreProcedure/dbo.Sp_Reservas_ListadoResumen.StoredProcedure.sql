USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 10/04/2026 | KPI global del listado general de reservas (pendientes, pagadas y saldo total) sin paginacion.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_ListadoResumen]
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL,
    @EstadosCsv NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @EstadosNormalizados NVARCHAR(200);
        SET @EstadosNormalizados = NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(@EstadosCsv)), N' ', N''), N';', N','), N'');

        IF OBJECT_ID('tempdb..#ReservasFiltradasResumen') IS NOT NULL
            DROP TABLE #ReservasFiltradasResumen;

        SELECT
            r.Estado,
            r.Total,
            r.Adelanto
        INTO #ReservasFiltradasResumen
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@FechaDesde IS NULL OR r.Fecha >= @FechaDesde)
          AND (@FechaHasta IS NULL OR r.Fecha <= @FechaHasta)
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND
          (
              (@Estado IS NOT NULL AND ((@Estado = 4 AND r.Estado IN (3, 4)) OR (@Estado <> 4 AND r.Estado = @Estado)))
              OR
              (
                  @Estado IS NULL
                  AND
                  (
                      @EstadosNormalizados IS NULL
                      OR EXISTS
                      (
                          SELECT 1
                          FROM STRING_SPLIT(@EstadosNormalizados, N',') estados
                          WHERE (TRY_CAST(estados.value AS INT) = 4 AND r.Estado IN (3, 4)) OR (TRY_CAST(estados.value AS INT) <> 4 AND TRY_CAST(estados.value AS INT) = r.Estado)
                      )
                  )
              )
          );

        SELECT
            SUM(CASE WHEN rf.Estado = 1 THEN 1 ELSE 0 END) AS TotalPendientes,
            SUM(CASE WHEN rf.Estado IN (3, 4) THEN 1 ELSE 0 END) AS TotalPagadas,
            CAST(SUM(rf.Total - rf.Adelanto) AS DECIMAL(18,2)) AS SaldoTotal
        FROM #ReservasFiltradasResumen rf;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
