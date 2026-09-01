-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Obtiene la cabecera y los tramos de una regla maestra de cuentas destino.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerCuentaDestinoMaestro
    @IdCuentaDestinoReglaMaestro INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            regla.IdCuentaDestinoReglaMaestro,
            regla.CodigoCuentaOrigen,
            cuenta.NombreCuenta AS NombreCuentaOrigen,
            regla.Activo,
            regla.Observacion
        FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = regla.CodigoCuentaOrigen
        WHERE regla.IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro;

        SELECT
            detalle.IdCuentaDestinoReglaDetalleMaestro,
            detalle.Orden,
            detalle.CodigoCuentaDestinoCargo,
            cargo.NombreCuenta AS NombreCuentaDestinoCargo,
            detalle.CodigoCuentaDestinoAbono,
            abono.NombreCuenta AS NombreCuentaDestinoAbono,
            detalle.Porcentaje,
            detalle.Activo
        FROM dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cargo
            ON cargo.CodigoCuenta = detalle.CodigoCuentaDestinoCargo
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS abono
            ON abono.CodigoCuenta = detalle.CodigoCuentaDestinoAbono
        WHERE detalle.IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro
        ORDER BY detalle.Orden;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
