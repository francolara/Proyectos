-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Elimina una cuenta maestra solo cuando no tiene hijos ni referencias de configuracion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarPlanCuentaMaestro
    @IdPlanCuentaMaestro INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoCuenta VARCHAR(20);
        SELECT @CodigoCuenta = CodigoCuenta FROM dbo.CON_PlanCuentaMaestro WHERE IdPlanCuentaMaestro = @IdPlanCuentaMaestro;

        IF @CodigoCuenta IS NULL
            RAISERROR(N'La cuenta maestra indicada no existe.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.CON_PlanCuentaMaestro WHERE CodigoCuentaPadre = @CodigoCuenta)
            RAISERROR(N'No se puede eliminar porque la cuenta tiene cuentas hijas.', 16, 1);

        IF EXISTS
        (
            SELECT 1 FROM dbo.CON_CuentaDestinoReglaMaestro WHERE CodigoCuentaOrigen = @CodigoCuenta
            UNION ALL
            SELECT 1 FROM dbo.CON_CuentaDestinoReglaDetalleMaestro WHERE CodigoCuentaDestinoCargo = @CodigoCuenta OR CodigoCuentaDestinoAbono = @CodigoCuenta
            UNION ALL
            SELECT 1 FROM dbo.CON_TipoImpuesto WHERE CodigoCuenta = @CodigoCuenta
            UNION ALL
            SELECT 1 FROM dbo.ADM_TipoComprobante
            WHERE CodigoCuentaVentaSoles = @CodigoCuenta OR CodigoCuentaVentaDolares = @CodigoCuenta
               OR CodigoCuentaCompraSoles = @CodigoCuenta OR CodigoCuentaCompraDolares = @CodigoCuenta
            UNION ALL
            SELECT 1 FROM dbo.ADM_ParametroMaestro
            WHERE CodigoParametro IN
            (
                'CUENTAGANANCIA', 'CUENTAGANANCIA_DC', 'CUENTAGANANCIA_AJ',
                'CUENTAPERDIDA', 'CUENTAPERDIDA_DC', 'CUENTAPERDIDA_AJ',
                'CTARETENCION', 'CTA_DEBE_CONSUMO', 'CTA_HABER_CONSUMO',
                'CTADETRACCION', 'CTADEPERCEPCION'
            )
              AND ValorParametro = @CodigoCuenta
        )
            RAISERROR(N'No se puede eliminar porque la cuenta esta utilizada en una configuracion maestra.', 16, 1);

        DELETE FROM dbo.CON_PlanCuentaMaestro WHERE IdPlanCuentaMaestro = @IdPlanCuentaMaestro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
