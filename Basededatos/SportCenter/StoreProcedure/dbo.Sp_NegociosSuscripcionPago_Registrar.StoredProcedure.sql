-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Registra cobro manual de suscripcion por negocio contemplando tipo y estado de pago.
-- Firma:         10/06/2026 | Soporta conciliacion y aplicacion automatica a la suscripcion cuando el cobro se confirma.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcionPago_Registrar
    @NegocioId INT,
    @TipoPago NVARCHAR(30),
    @EstadoPago NVARCHAR(20) = N'PAGADO',
    @Monto DECIMAL(12,2),
    @Moneda NVARCHAR(10) = N'PEN',
    @FechaPago DATETIME2(7) = NULL,
    @FechaVencimiento DATE = NULL,
    @OperacionNumero NVARCHAR(100) = NULL,
    @EntidadFinanciera NVARCHAR(120) = NULL,
    @ReferenciaExterna NVARCHAR(120) = NULL,
    @AccionAplicacion NVARCHAR(30) = NULL,
    @AplicarAlConfirmar BIT = 0,
    @TipoCobroObjetivo NVARCHAR(20) = NULL,
    @FechaInicioPlanObjetivo DATE = NULL,
    @DiasGraciaObjetivo INT = NULL,
    @Observacion NVARCHAR(500) = NULL,
    @NegocioSuscripcionMovimientoId INT = NULL,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @NegocioSuscripcionId INT;
        DECLARE @PagoId INT;
        DECLARE @TipoPagoNorm NVARCHAR(30) = UPPER(LTRIM(RTRIM(COALESCE(@TipoPago, N''))));
        DECLARE @EstadoPagoNorm NVARCHAR(20) = UPPER(LTRIM(RTRIM(COALESCE(@EstadoPago, N'PAGADO'))));
        DECLARE @MonedaNorm NVARCHAR(10) = UPPER(LTRIM(RTRIM(COALESCE(@Moneda, N'PEN'))));
        DECLARE @AccionAplicacionNorm NVARCHAR(30) = UPPER(LTRIM(RTRIM(COALESCE(@AccionAplicacion, N''))));
        DECLARE @TipoCobroObjetivoNorm NVARCHAR(20) = UPPER(LTRIM(RTRIM(COALESCE(@TipoCobroObjetivo, N''))));
        DECLARE @FechaFinPlanObjetivoCalculada DATE;
        DECLARE @DiasGraciaObjetivoFinal INT;

        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        IF @TipoPagoNorm NOT IN (N'EFECTIVO', N'TRANSFERENCIA', N'YAPE', N'PLIN', N'LINK_PAGO', N'PASARELA')
            RAISERROR('Tipo de pago invalido.', 16, 1);

        IF @EstadoPagoNorm NOT IN (N'PENDIENTE', N'PAGADO', N'OBSERVADO', N'ANULADO')
            RAISERROR('Estado de pago invalido.', 16, 1);

        IF @Monto IS NULL OR @Monto <= 0
            RAISERROR('El monto del cobro debe ser mayor a cero.', 16, 1);

        IF NULLIF(@AccionAplicacionNorm, N'') IS NOT NULL
           AND @AccionAplicacionNorm NOT IN (N'ACTIVACION_CONTRATO', N'RENOVACION', N'CAMBIO_PLAN')
            RAISERROR('Accion de aplicacion invalida.', 16, 1);

        IF @AplicarAlConfirmar = 1 AND NULLIF(@AccionAplicacionNorm, N'') IS NULL
            RAISERROR('Selecciona una accion de aplicacion para el cobro conciliable.', 16, 1);

        IF @AccionAplicacionNorm IN (N'ACTIVACION_CONTRATO', N'CAMBIO_PLAN')
        BEGIN
            IF @TipoCobroObjetivoNorm NOT IN (N'MENSUAL', N'TRIMESTRAL', N'SEMESTRAL', N'ANUAL')
                RAISERROR('Tipo de cobro objetivo invalido para la aplicacion del contrato.', 16, 1);

            IF @FechaInicioPlanObjetivo IS NULL
                RAISERROR('La fecha de inicio objetivo es obligatoria para la aplicacion del contrato.', 16, 1);

            IF @DiasGraciaObjetivo IS NULL OR @DiasGraciaObjetivo < 0
                SET @DiasGraciaObjetivo = 5;

            SET @DiasGraciaObjetivoFinal = COALESCE(@DiasGraciaObjetivo, 5);

            SET @FechaFinPlanObjetivoCalculada = CASE @TipoCobroObjetivoNorm
                WHEN N'TRIMESTRAL' THEN DATEADD(MONTH, 3, @FechaInicioPlanObjetivo)
                WHEN N'SEMESTRAL' THEN DATEADD(MONTH, 6, @FechaInicioPlanObjetivo)
                WHEN N'ANUAL' THEN DATEADD(YEAR, 1, @FechaInicioPlanObjetivo)
                ELSE DATEADD(MONTH, 1, @FechaInicioPlanObjetivo)
            END;
        END
        ELSE
        BEGIN
            SET @DiasGraciaObjetivoFinal = COALESCE(@DiasGraciaObjetivo, 5);
        END

        SELECT
            @NegocioSuscripcionId = ns.Id
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @NegocioSuscripcionId IS NULL
            RAISERROR('El negocio no tiene una suscripcion registrada.', 16, 1);

        IF @NegocioSuscripcionMovimientoId IS NULL
        BEGIN
            SELECT TOP (1)
                @NegocioSuscripcionMovimientoId = m.Id
            FROM dbo.NegociosSuscripcionMovimiento m
            WHERE m.NegocioId = @NegocioId
            ORDER BY m.FechaCreacion DESC, m.Id DESC;
        END

        INSERT INTO dbo.NegociosSuscripcionPago
        (
            NegocioId, NegocioSuscripcionId, NegocioSuscripcionMovimientoId,
            TipoPago, EstadoPago, Monto, Moneda,
            FechaPago, FechaVencimiento, OperacionNumero, EntidadFinanciera,
            ReferenciaExterna, AccionAplicacion, AplicarAlConfirmar, AplicadoSuscripcion,
            FechaAplicacion, UsuarioAplicacion, TipoCobroObjetivo, FechaInicioPlanObjetivo,
            DiasGraciaObjetivo, Observacion, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @NegocioSuscripcionId, @NegocioSuscripcionMovimientoId,
            @TipoPagoNorm, @EstadoPagoNorm, @Monto, @MonedaNorm,
            COALESCE(@FechaPago, SYSUTCDATETIME()), @FechaVencimiento,
            NULLIF(LTRIM(RTRIM(COALESCE(@OperacionNumero, N''))), N''),
            NULLIF(LTRIM(RTRIM(COALESCE(@EntidadFinanciera, N''))), N''),
            NULLIF(LTRIM(RTRIM(COALESCE(@ReferenciaExterna, N''))), N''),
            NULLIF(@AccionAplicacionNorm, N''), @AplicarAlConfirmar, 0,
            NULL, NULL, NULLIF(@TipoCobroObjetivoNorm, N''), @FechaInicioPlanObjetivo,
            @DiasGraciaObjetivo,
            NULLIF(LTRIM(RTRIM(COALESCE(@Observacion, N''))), N''),
            SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
        );

        SET @PagoId = CAST(SCOPE_IDENTITY() AS INT);

        IF @EstadoPagoNorm = N'PAGADO' AND @AplicarAlConfirmar = 1 AND NULLIF(@AccionAplicacionNorm, N'') IS NOT NULL
        BEGIN
            IF @AccionAplicacionNorm = N'RENOVACION'
                EXEC dbo.Sp_NegociosSuscripcion_RenovarPlan @NegocioId = @NegocioId, @Usuario = @Usuario;
            ELSE IF @AccionAplicacionNorm = N'ACTIVACION_CONTRATO'
                EXEC dbo.Sp_NegociosSuscripcion_ActivarPlan
                    @NegocioId = @NegocioId,
                    @TipoCobro = @TipoCobroObjetivoNorm,
                    @FechaInicioPlan = @FechaInicioPlanObjetivo,
                    @FechaFinPlan = @FechaFinPlanObjetivoCalculada,
                    @DiasGracia = @DiasGraciaObjetivoFinal,
                    @Usuario = @Usuario;
            ELSE IF @AccionAplicacionNorm = N'CAMBIO_PLAN'
                EXEC dbo.Sp_NegociosSuscripcion_CambiarPlan
                    @NegocioId = @NegocioId,
                    @TipoCobro = @TipoCobroObjetivoNorm,
                    @FechaInicioPlan = @FechaInicioPlanObjetivo,
                    @FechaFinPlan = @FechaFinPlanObjetivoCalculada,
                    @DiasGracia = @DiasGraciaObjetivoFinal,
                    @Observacion = @Observacion,
                    @Usuario = @Usuario;

            UPDATE dbo.NegociosSuscripcionPago
            SET AplicadoSuscripcion = 1,
                FechaAplicacion = SYSUTCDATETIME(),
                UsuarioAplicacion = COALESCE(@Usuario, N'sistema'),
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE Id = @PagoId;
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
