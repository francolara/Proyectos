
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 18/06/2026 | Nuevo resumen comercial por fecha de pago para separar KPIs de cobranza de los KPIs operativos.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reportes_ResumenCobranza]
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            COUNT(1) AS CantidadPagos,
            COUNT(DISTINCT p.ReservaId) AS ReservasCobradas,
            CAST(COALESCE(SUM(p.Monto), 0) AS DECIMAL(18,2)) AS MontoCobrado
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND CAST(p.FechaPago AS DATE) >= @FechaDesde
          AND CAST(p.FechaPago AS DATE) <= @FechaHasta
          AND r.Estado <> 5;
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
