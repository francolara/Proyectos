
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 18/06/2026 | Nuevo reporte operativo por fecha de reserva para separar operacion y cobranza en dashboard y reportes.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reportes_ReservasPorDia]
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            r.Fecha,
            COUNT(1) AS CantidadReservas,
            CAST(COALESCE(SUM(r.Total), 0) AS DECIMAL(18,2)) AS MontoReservado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado <> 5
        GROUP BY r.Fecha
        ORDER BY r.Fecha ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
