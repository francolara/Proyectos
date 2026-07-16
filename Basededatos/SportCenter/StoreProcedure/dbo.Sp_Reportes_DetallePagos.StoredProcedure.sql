-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/07/2026
-- Description:   Detalle ejecutivo de pagos por fecha real de cobro para impresion de reportes.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_DetallePagos
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            p.Id AS PagoId,
            p.FechaPago,
            r.Id AS ReservaId,
            r.Fecha AS FechaReserva,
            r.HoraInicio,
            r.HoraFin,
            c.NombresORazonSocial AS Cliente,
            c.NumeroDocumento AS ClienteDocumento,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            fp.Nombre AS FormaPago,
            p.NumeroOperacion,
            CAST(p.Monto AS DECIMAL(18, 2)) AS Monto
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND CAST(p.FechaPago AS DATE) >= @FechaDesde
          AND CAST(p.FechaPago AS DATE) <= @FechaHasta
          AND r.Estado <> 5
        ORDER BY p.FechaPago ASC, p.Id ASC;
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
