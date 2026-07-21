-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Lista cobros de suscripcion por negocio devolviendo resumen acumulado, estado de conciliacion e historial reciente.
-- Firma:         FRANCO LARA - 21/07/2026 | Devuelve el plan comercial y limites objetivo aplicados con cada cobro.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcionPago_ListarPorNegocio
    @NegocioId INT,
    @Top INT = 8,
    @CantidadPagos INT OUTPUT,
    @MontoTotalPagado DECIMAL(12,2) OUTPUT,
    @UltimaFechaPago DATETIME2(7) OUTPUT,
    @UltimoMonto DECIMAL(12,2) OUTPUT,
    @UltimoTipoPago NVARCHAR(30) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        SET @Top = CASE WHEN ISNULL(@Top, 0) <= 0 THEN 8 ELSE @Top END;

        SELECT
            @CantidadPagos = COUNT(1),
            @MontoTotalPagado = COALESCE(SUM(CASE WHEN p.EstadoPago = N'PAGADO' THEN p.Monto ELSE 0 END), 0)
        FROM dbo.NegociosSuscripcionPago p
        WHERE p.NegocioId = @NegocioId
          AND p.EstadoPago <> N'ANULADO';

        SELECT TOP (1)
            @UltimaFechaPago = p.FechaPago,
            @UltimoMonto = p.Monto,
            @UltimoTipoPago = p.TipoPago
        FROM dbo.NegociosSuscripcionPago p
        WHERE p.NegocioId = @NegocioId
          AND p.EstadoPago <> N'ANULADO'
        ORDER BY p.FechaPago DESC, p.Id DESC;

        SELECT TOP (@Top)
            p.Id,
            p.TipoPago,
            p.EstadoPago,
            p.Monto,
            p.Moneda,
            p.FechaPago,
            p.FechaVencimiento,
            p.OperacionNumero,
            p.EntidadFinanciera,
            p.ReferenciaExterna,
            p.Observacion,
            p.FechaCreacion,
            p.UsuarioCreacion,
            m.TipoMovimiento,
            p.AccionAplicacion,
            p.AplicarAlConfirmar,
            p.AplicadoSuscripcion,
            p.FechaAplicacion,
            p.UsuarioAplicacion,
            p.TipoCobroObjetivo,
            p.FechaInicioPlanObjetivo,
            p.DiasGraciaObjetivo,
            p.PlanComercialObjetivo,
            p.TipoPlanObjetivo,
            p.SedesPermitidasObjetivo,
            p.EspaciosPermitidosObjetivo,
            p.UsuariosPermitidosObjetivo
        FROM dbo.NegociosSuscripcionPago p
        LEFT JOIN dbo.NegociosSuscripcionMovimiento m ON m.Id = p.NegocioSuscripcionMovimientoId
        WHERE p.NegocioId = @NegocioId
        ORDER BY p.FechaPago DESC, p.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
