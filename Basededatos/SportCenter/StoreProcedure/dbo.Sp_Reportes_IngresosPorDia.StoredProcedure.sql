
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 13/04/2026 | Normaliza script a CREATE OR ALTER, mantiene filtro por sede y excluye canceladas (Estado=5) del conteo/monto KPI diario.
-- Firma: Codex - 18/06/2026 | Reorienta el reporte de ingresos a fecha de pago para que cobros adelantados y parciales se agrupen por el dia real de cobranza.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reportes_IngresosPorDia]
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        ;WITH PagosBase AS
        (
            SELECT
                CAST(p.FechaPago AS DATE) AS FechaPago,
                p.ReservaId,
                p.Monto
            FROM dbo.Pagos p
            INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND CAST(p.FechaPago AS DATE) >= @FechaDesde
              AND CAST(p.FechaPago AS DATE) <= @FechaHasta
              AND r.Estado <> 5
        )
        SELECT
            pb.FechaPago AS Fecha,
            COUNT(DISTINCT pb.ReservaId) AS CantidadReservas,
            COALESCE(SUM(pb.Monto), 0) AS Ingresos
        FROM PagosBase pb
        GROUP BY pb.FechaPago
        ORDER BY pb.FechaPago ASC;
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
