-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Confirma y aplica un cobro pendiente de suscripcion por cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ConfirmarPagoSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @IdCuentaAdministradoraSuscripcionPago INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @AccionAplicacion NVARCHAR(30);
        DECLARE @AplicarAlConfirmar BIT;
        DECLARE @AplicadoSuscripcion BIT;
        DECLARE @TipoCobroObjetivo NVARCHAR(20);
        DECLARE @FechaInicioPlanObjetivo DATE;
        DECLARE @DiasGraciaObjetivo INT;
        DECLARE @Observacion NVARCHAR(500);
        DECLARE @FechaFinPlanObjetivo DATE;
        DECLARE @DiasGraciaNormalizado INT;
        DECLARE @TipoCobroContrato NVARCHAR(20);
        DECLARE @ObservacionContrato NVARCHAR(500);

        SELECT
            @AccionAplicacion = p.AccionAplicacion,
            @AplicarAlConfirmar = p.AplicarAlConfirmar,
            @AplicadoSuscripcion = p.AplicadoSuscripcion,
            @TipoCobroObjetivo = p.TipoCobroObjetivo,
            @FechaInicioPlanObjetivo = p.FechaInicioPlanObjetivo,
            @DiasGraciaObjetivo = p.DiasGraciaObjetivo,
            @Observacion = p.Observacion
        FROM dbo.SEG_CuentaAdministradoraSuscripcionPago AS p
        WHERE p.IdCuentaAdministradoraSuscripcionPago = @IdCuentaAdministradoraSuscripcionPago
          AND p.IdCuentaAdministradora = @IdCuentaAdministradora;

        IF @AccionAplicacion IS NULL AND @AplicarAlConfirmar IS NULL
        BEGIN
            RAISERROR (N'No se encontro el cobro de suscripcion seleccionado.', 16, 1);
            RETURN;
        END;

        UPDATE dbo.SEG_CuentaAdministradoraSuscripcionPago
        SET EstadoPago = N'PAGADO',
            EstadoPasarela = COALESCE(EstadoPasarela, N'CONFIRMADO'),
            FechaConfirmacionPasarela = COALESCE(FechaConfirmacionPasarela, SYSDATETIME()),
            FechaActualizacion = SYSDATETIME(),
            UsuarioActualizacion = @UsuarioRegistro
        WHERE IdCuentaAdministradoraSuscripcionPago = @IdCuentaAdministradoraSuscripcionPago;

        IF @AplicarAlConfirmar = 1
           AND ISNULL(@AplicadoSuscripcion, 0) = 0
           AND UPPER(LTRIM(RTRIM(ISNULL(@AccionAplicacion, N'')))) = N'ACTIVAR_CONTRATO'
           AND @FechaInicioPlanObjetivo IS NOT NULL
        BEGIN
            SET @DiasGraciaNormalizado = CASE WHEN @DiasGraciaObjetivo IS NULL OR @DiasGraciaObjetivo < 0 THEN 5 ELSE @DiasGraciaObjetivo END;
            SET @TipoCobroContrato = COALESCE(NULLIF(LTRIM(RTRIM(@TipoCobroObjetivo)), N''), N'MENSUAL');
            SET @ObservacionContrato = COALESCE(NULLIF(LTRIM(RTRIM(@Observacion)), N''), N'Cobro confirmado y aplicado al contrato.');

            SET @FechaFinPlanObjetivo =
                CASE UPPER(LTRIM(RTRIM(@TipoCobroContrato)))
                    WHEN N'TRIMESTRAL' THEN DATEADD(DAY, -1, DATEADD(MONTH, 3, @FechaInicioPlanObjetivo))
                    WHEN N'SEMESTRAL' THEN DATEADD(DAY, -1, DATEADD(MONTH, 6, @FechaInicioPlanObjetivo))
                    WHEN N'ANUAL' THEN DATEADD(DAY, -1, DATEADD(YEAR, 1, @FechaInicioPlanObjetivo))
                    ELSE DATEADD(DAY, -1, DATEADD(MONTH, 1, @FechaInicioPlanObjetivo))
                END;

            EXEC dbo.usp_SEG_ActivarContratoCuentaAdministradora
                @IdCuentaAdministradora = @IdCuentaAdministradora,
                @TipoCobro = @TipoCobroContrato,
                @FechaInicioPlan = @FechaInicioPlanObjetivo,
                @FechaFinPlan = @FechaFinPlanObjetivo,
                @DiasGracia = @DiasGraciaNormalizado,
                @Observacion = @ObservacionContrato,
                @UsuarioRegistro = @UsuarioRegistro;

            UPDATE dbo.SEG_CuentaAdministradoraSuscripcionPago
            SET AplicadoSuscripcion = 1,
                FechaAplicacion = SYSDATETIME(),
                UsuarioAplicacion = @UsuarioRegistro,
                FechaActualizacion = SYSDATETIME(),
                UsuarioActualizacion = @UsuarioRegistro
            WHERE IdCuentaAdministradoraSuscripcionPago = @IdCuentaAdministradoraSuscripcionPago;
        END;

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
