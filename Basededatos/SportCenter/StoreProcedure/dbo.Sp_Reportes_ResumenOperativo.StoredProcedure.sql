USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 13/04/2026 | Resumen operativo de Reportes por rango/sede con exclusión de reservas canceladas (Estado=5) en total de reservas y KPIs de monto/saldo.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reportes_ResumenOperativo]
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        ;WITH ReservasBase AS
        (
            SELECT
                r.Id,
                r.Estado,
                r.Total
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND r.Fecha >= @FechaDesde
              AND r.Fecha <= @FechaHasta
        ),
        PagosPorReserva AS
        (
            SELECT
                p.ReservaId,
                SUM(p.Monto) AS MontoCobrado
            FROM dbo.Pagos p
            INNER JOIN ReservasBase rb ON rb.Id = p.ReservaId
            GROUP BY p.ReservaId
        )
        SELECT
            COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN 1 ELSE 0 END), 0) AS TotalReservas,
            COALESCE(SUM(CASE WHEN rb.Estado = 1 THEN 1 ELSE 0 END), 0) AS TotalPendientes,
            COALESCE(SUM(CASE WHEN rb.Estado IN (2, 3) THEN 1 ELSE 0 END), 0) AS TotalConfirmadas,
            COALESCE(SUM(CASE WHEN rb.Estado = 4 THEN 1 ELSE 0 END), 0) AS TotalPagadas,
            COALESCE(SUM(CASE WHEN rb.Estado = 5 THEN 1 ELSE 0 END), 0) AS TotalCanceladas,
            COALESCE(SUM(CASE WHEN rb.Estado = 6 THEN 1 ELSE 0 END), 0) AS TotalNoShow,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN rb.Total ELSE 0 END), 0) AS DECIMAL(18,2)) AS MontoReservado,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN COALESCE(pr.MontoCobrado, 0) ELSE 0 END), 0) AS DECIMAL(18,2)) AS MontoCobrado,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 AND rb.Total - COALESCE(pr.MontoCobrado, 0) > 0 THEN rb.Total - COALESCE(pr.MontoCobrado, 0) ELSE 0 END), 0) AS DECIMAL(18,2)) AS SaldoPendiente
        FROM ReservasBase rb
        LEFT JOIN PagosPorReserva pr ON pr.ReservaId = rb.Id;
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
