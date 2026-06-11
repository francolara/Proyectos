-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Confirma un cobro pendiente de suscripcion y aplica automaticamente la accion comercial configurada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcionPago_ConfirmarAplicar
    @NegocioId INT,
    @PagoId INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @EstadoPago NVARCHAR(20);
        DECLARE @AplicarAlConfirmar BIT;
        DECLARE @AplicadoSuscripcion BIT;
        DECLARE @AccionAplicacion NVARCHAR(30);
        DECLARE @TipoCobroObjetivo NVARCHAR(20);
        DECLARE @FechaInicioPlanObjetivo DATE;
        DECLARE @DiasGraciaObjetivo INT;
        DECLARE @FechaFinPlanObjetivoCalculada DATE;
        DECLARE @DiasGraciaObjetivoFinal INT;
        DECLARE @ObservacionAplicacion NVARCHAR(500) = N'Aplicacion automatica desde confirmacion de cobro.';

        IF @NegocioId IS NULL OR @NegocioId <= 0 OR @PagoId IS NULL OR @PagoId <= 0
            RAISERROR('Cobro invalido.', 16, 1);

        SELECT
            @EstadoPago = UPPER(LTRIM(RTRIM(COALESCE(p.EstadoPago, N'')))),
            @AplicarAlConfirmar = COALESCE(p.AplicarAlConfirmar, 0),
            @AplicadoSuscripcion = COALESCE(p.AplicadoSuscripcion, 0),
            @AccionAplicacion = UPPER(LTRIM(RTRIM(COALESCE(p.AccionAplicacion, N'')))),
            @TipoCobroObjetivo = UPPER(LTRIM(RTRIM(COALESCE(p.TipoCobroObjetivo, N'')))),
            @FechaInicioPlanObjetivo = p.FechaInicioPlanObjetivo,
            @DiasGraciaObjetivo = COALESCE(p.DiasGraciaObjetivo, 5)
        FROM dbo.NegociosSuscripcionPago p
        WHERE p.Id = @PagoId
          AND p.NegocioId = @NegocioId;

        IF @EstadoPago IS NULL
            RAISERROR('No se encontro el cobro seleccionado.', 16, 1);

        IF @EstadoPago = N'ANULADO'
            RAISERROR('No se puede confirmar un cobro anulado.', 16, 1);

        UPDATE dbo.NegociosSuscripcionPago
        SET EstadoPago = N'PAGADO',
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        WHERE Id = @PagoId
          AND NegocioId = @NegocioId;

        IF COALESCE(@AplicarAlConfirmar, 0) = 1 AND COALESCE(@AplicadoSuscripcion, 0) = 0 AND NULLIF(@AccionAplicacion, N'') IS NOT NULL
        BEGIN
            IF @AccionAplicacion IN (N'ACTIVACION_CONTRATO', N'CAMBIO_PLAN')
            BEGIN
                SET @DiasGraciaObjetivoFinal = COALESCE(@DiasGraciaObjetivo, 5);
                SET @FechaFinPlanObjetivoCalculada = CASE @TipoCobroObjetivo
                    WHEN N'TRIMESTRAL' THEN DATEADD(MONTH, 3, @FechaInicioPlanObjetivo)
                    WHEN N'SEMESTRAL' THEN DATEADD(MONTH, 6, @FechaInicioPlanObjetivo)
                    WHEN N'ANUAL' THEN DATEADD(YEAR, 1, @FechaInicioPlanObjetivo)
                    ELSE DATEADD(MONTH, 1, @FechaInicioPlanObjetivo)
                END;
            END;
            
            IF @AccionAplicacion = N'RENOVACION'
                SET @DiasGraciaObjetivoFinal = COALESCE(@DiasGraciaObjetivo, 5);

            IF @AccionAplicacion = N'RENOVACION'
                EXEC dbo.Sp_NegociosSuscripcion_RenovarPlan @NegocioId = @NegocioId, @Usuario = @Usuario;
            ELSE IF @AccionAplicacion = N'ACTIVACION_CONTRATO'
                EXEC dbo.Sp_NegociosSuscripcion_ActivarPlan
                    @NegocioId = @NegocioId,
                    @TipoCobro = @TipoCobroObjetivo,
                    @FechaInicioPlan = @FechaInicioPlanObjetivo,
                    @FechaFinPlan = @FechaFinPlanObjetivoCalculada,
                    @DiasGracia = @DiasGraciaObjetivoFinal,
                    @Usuario = @Usuario;
            ELSE IF @AccionAplicacion = N'CAMBIO_PLAN'
                EXEC dbo.Sp_NegociosSuscripcion_CambiarPlan
                    @NegocioId = @NegocioId,
                    @TipoCobro = @TipoCobroObjetivo,
                    @FechaInicioPlan = @FechaInicioPlanObjetivo,
                    @FechaFinPlan = @FechaFinPlanObjetivoCalculada,
                    @DiasGracia = @DiasGraciaObjetivoFinal,
                    @Observacion = @ObservacionAplicacion,
                    @Usuario = @Usuario;

            UPDATE dbo.NegociosSuscripcionPago
            SET AplicadoSuscripcion = 1,
                FechaAplicacion = SYSUTCDATETIME(),
                UsuarioAplicacion = COALESCE(@Usuario, N'sistema'),
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE Id = @PagoId
              AND NegocioId = @NegocioId;
        END
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
