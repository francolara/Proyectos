-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/07/2026
-- Description:   Detalle ejecutivo de reservas, importes y saldos para impresion de reportes.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_DetalleReservas
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        ;WITH PagosPorReserva AS
        (
            SELECT
                p.ReservaId,
                CAST(SUM(p.Monto) AS DECIMAL(18, 2)) AS MontoPagado
            FROM dbo.Pagos p
            GROUP BY p.ReservaId
        ),
        CuponPorReserva AS
        (
            SELECT
                cu.ReservaId,
                MAX(c.CodigoCupon) AS CodigoCupon,
                CAST(SUM(cu.MontoDescuento) AS DECIMAL(18, 2)) AS MontoDescuento
            FROM dbo.CuponesUso cu
            INNER JOIN dbo.Cupones c ON c.Id = cu.CuponId
            GROUP BY cu.ReservaId
        )
        SELECT
            r.Id AS ReservaId,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            c.NombresORazonSocial AS Cliente,
            c.NumeroDocumento AS ClienteDocumento,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Estado AS EstadoCodigo,
            CASE r.Estado
                WHEN 1 THEN N'Pendiente'
                WHEN 2 THEN N'Confirmada'
                WHEN 3 THEN N'En uso'
                WHEN 4 THEN N'Pagada'
                WHEN 5 THEN N'Cancelada'
                WHEN 6 THEN N'No asistio'
                ELSE N'Sin estado'
            END AS Estado,
            r.CanalOrigen,
            CAST(r.Total AS DECIMAL(18, 2)) AS Total,
            COALESCE(cr.MontoDescuento, 0) AS Descuento,
            COALESCE(pr.MontoPagado, 0) AS MontoPagado,
            CAST(CASE
                WHEN r.Estado = 5 OR r.Total <= COALESCE(pr.MontoPagado, 0) THEN 0
                ELSE r.Total - COALESCE(pr.MontoPagado, 0)
            END AS DECIMAL(18, 2)) AS SaldoPendiente,
            cr.CodigoCupon
        FROM dbo.Reservas r
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN PagosPorReserva pr ON pr.ReservaId = r.Id
        LEFT JOIN CuponPorReserva cr ON cr.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
        ORDER BY r.Fecha ASC, r.HoraInicio ASC, r.Id ASC;
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
