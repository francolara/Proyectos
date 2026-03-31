-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Filtros de reservas por rango/sede/espacio/estado para vista calendario semanal.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/03/2026
-- Description:   Extiende Sp_Reservas_Listar para soportar filtro de estados multiples (@EstadosCsv) sin romper @Estado.
-- Firma:         Codex - 30/03/2026
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Listar
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

        SELECT TOP (300)
            r.Id,
            c.NombresORazonSocial AS Cliente,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            CAST(r.Estado AS NVARCHAR(20)) AS Estado
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
              (@Estado IS NOT NULL AND r.Estado = @Estado)
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
                          WHERE TRY_CAST(estados.value AS INT) = r.Estado
                      )
                  )
              )
          )
        ORDER BY r.Fecha ASC, r.HoraInicio ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
