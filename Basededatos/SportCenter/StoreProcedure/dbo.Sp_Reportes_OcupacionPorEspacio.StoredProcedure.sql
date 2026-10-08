
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 13/04/2026 | Incluye SedeId y EspacioDeportivoId para drill-down desde Reportes a Reservas/Pagos.
-- Firma: FRANCO LARA - 06/10/2026 | Separa ocupacion e importes por CodigoMoneda canonico seleccionado.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reportes_OcupacionPorEspacio]
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @CodigoMoneda NVARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        ;WITH PagosPorReserva AS
        (
            SELECT
                p.ReservaId,
                SUM(p.Monto) AS MontoCobrado
            FROM dbo.Pagos p
            WHERE p.CodigoMoneda = @CodigoMoneda
            GROUP BY p.ReservaId
        )
        SELECT
            s.Id AS SedeId,
            e.Id AS EspacioDeportivoId,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            COUNT(1) AS CantidadReservas,
            CAST(SUM(DATEDIFF(MINUTE, r.HoraInicio, r.HoraFin)) / 60.0 AS DECIMAL(10,2)) AS HorasReservadas,
            SUM(r.Total) AS MontoReservado,
            SUM(COALESCE(pr.MontoCobrado, 0)) AS MontoCobrado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN PagosPorReserva pr ON pr.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.CodigoMoneda = @CodigoMoneda
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado NOT IN (5, 6)
        GROUP BY s.Id, e.Id, s.Nombre, e.Nombre
        ORDER BY HorasReservadas DESC, CantidadReservas DESC;
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
