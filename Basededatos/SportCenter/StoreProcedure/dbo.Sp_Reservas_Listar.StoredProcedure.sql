USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Listar]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 34_Clientes_NombreEquipo_Reservas.sql (linea 248)
-- Firma: Codex - 05/04/2026 | Filtro de estado Pagada incluye estados historicos 3 y 4; retiro operativo de En uso.
-- Firma: Codex - 07/04/2026 | Incluye Adelanto, SaldoPendiente, paginacion backend con total de registros para listado general, y separa Cliente/Equipo en columnas independientes.
CREATE OR ALTER  PROCEDURE [dbo].[Sp_Reservas_Listar]
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL,
    @EstadosCsv NVARCHAR(200) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @EstadosNormalizados NVARCHAR(200);
        SET @EstadosNormalizados = NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(@EstadosCsv)), N' ', N''), N';', N','), N'');
        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;

        IF OBJECT_ID('tempdb..#ReservasFiltradas') IS NOT NULL
            DROP TABLE #ReservasFiltradas;

        SELECT
            r.Id,
            CAST(c.NombresORazonSocial AS NVARCHAR(250)) AS Cliente,
            CAST(NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') AS NVARCHAR(120)) AS Equipo,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.Adelanto,
            (r.Total - r.Adelanto) AS SaldoPendiente,
            CAST(r.Estado AS NVARCHAR(20)) AS Estado
        INTO #ReservasFiltradas
        FROM dbo.Reservas r
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
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
            @TotalRegistros = COUNT(1)
        FROM #ReservasFiltradas;

        SELECT
            r.Id,
            r.Cliente,
            r.Equipo,
            r.Espacio,
            r.Sede,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.Adelanto,
            r.SaldoPendiente,
            r.Estado
        FROM #ReservasFiltradas r
        ORDER BY r.Fecha ASC, r.HoraInicio ASC, r.Id ASC
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
