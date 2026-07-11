-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Lista los cobros de suscripcion de una cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarPagosSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @Top INT = 20
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT TOP (CASE WHEN @Top <= 0 THEN 20 ELSE @Top END)
            p.IdCuentaAdministradoraSuscripcionPago,
            p.TipoPago,
            p.EstadoPago,
            p.Monto,
            p.Moneda,
            p.FechaPago,
            p.FechaVencimiento,
            p.OperacionNumero,
            p.EntidadFinanciera,
            p.ReferenciaExterna,
            p.ProveedorPasarela,
            p.TransaccionPasarelaId,
            p.PagoPasarelaId,
            p.EstadoPasarela,
            p.AccionAplicacion,
            p.AplicarAlConfirmar,
            p.AplicadoSuscripcion,
            p.FechaAplicacion,
            p.UsuarioAplicacion,
            p.TipoCobroObjetivo,
            p.FechaInicioPlanObjetivo,
            p.DiasGraciaObjetivo,
            p.Observacion,
            p.FechaRegistro,
            p.UsuarioRegistro
        FROM dbo.SEG_CuentaAdministradoraSuscripcionPago AS p
        WHERE p.IdCuentaAdministradora = @IdCuentaAdministradora
        ORDER BY p.FechaPago DESC, p.IdCuentaAdministradoraSuscripcionPago DESC;

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
