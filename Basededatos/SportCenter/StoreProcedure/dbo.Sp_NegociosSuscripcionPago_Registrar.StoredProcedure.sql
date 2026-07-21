-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Registra cobro manual de suscripcion por negocio contemplando tipo y estado de pago.
-- Firma:         10/06/2026 | Soporta conciliacion y aplicacion automatica a la suscripcion cuando el cobro se confirma.
-- Firma:         FRANCO LARA - 21/07/2026 | Aplica al guardar el plan comercial, contrato y limites, conservando la fotografia anterior y nueva en el historial.
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
    @PlanComercialObjetivo NVARCHAR(20) = NULL,
    @SedesPermitidasObjetivo INT = NULL,
    @EspaciosPermitidosObjetivo INT = NULL,
    @UsuariosPermitidosObjetivo INT = NULL,
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
        DECLARE @PlanComercialObjetivoNorm NVARCHAR(20) = UPPER(LTRIM(RTRIM(COALESCE(@PlanComercialObjetivo, N''))));
        DECLARE @TipoPlanObjetivo NVARCHAR(20);
        DECLARE @FechaFinPlanObjetivoCalculada DATE;
        DECLARE @DiasGraciaObjetivoFinal INT;
        DECLARE @PlanComercialAnterior NVARCHAR(20);
        DECLARE @TipoPlanAnterior NVARCHAR(20);
        DECLARE @SedesPermitidasAnterior INT;
        DECLARE @EspaciosPermitidosAnterior INT;
        DECLARE @UsuariosPermitidosAnterior INT;
        DECLARE @MovimientoAplicadoId INT;

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

        IF NULLIF(@AccionAplicacionNorm, N'') IS NOT NULL
        BEGIN
            IF @PlanComercialObjetivoNorm NOT IN (N'ESENCIAL', N'PRO')
                RAISERROR('Plan comercial objetivo invalido. Usa ESENCIAL o PRO.', 16, 1);

            IF COALESCE(@SedesPermitidasObjetivo, 0) < 1
               OR COALESCE(@EspaciosPermitidosObjetivo, 0) < 1
               OR COALESCE(@UsuariosPermitidosObjetivo, 0) < 1
                RAISERROR('Los limites objetivo deben ser mayores a cero.', 16, 1);

            SET @TipoPlanObjetivo = CASE WHEN @PlanComercialObjetivoNorm = N'PRO' THEN N'Full' ELSE N'Basico' END;
        END;

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
            @NegocioSuscripcionId = ns.Id,
            @PlanComercialAnterior = COALESCE(NULLIF(ns.PlanComercial, N''), CASE WHEN ns.EsPrueba = 1 THEN N'PRUEBA' ELSE N'ESENCIAL' END)
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @NegocioSuscripcionId IS NULL
            RAISERROR('El negocio no tiene una suscripcion registrada.', 16, 1);

        SELECT
            @TipoPlanAnterior = COALESCE(NULLIF(n.TipoPlan, N''), N'Basico'),
            @SedesPermitidasAnterior = COALESCE(n.SedesPermitidas, 1),
            @EspaciosPermitidosAnterior = COALESCE(n.EspaciosPermitidos, 1),
            @UsuariosPermitidosAnterior = COALESCE(n.UsuariosPermitidos, 1)
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId;

        IF @NegocioSuscripcionMovimientoId IS NULL
        BEGIN
            SELECT TOP (1)
                @NegocioSuscripcionMovimientoId = m.Id
            FROM dbo.NegociosSuscripcionMovimiento m
            WHERE m.NegocioId = @NegocioId
            ORDER BY m.FechaCreacion DESC, m.Id DESC;
        END

        BEGIN TRANSACTION;

        INSERT INTO dbo.NegociosSuscripcionPago
        (
            NegocioId, NegocioSuscripcionId, NegocioSuscripcionMovimientoId,
            TipoPago, EstadoPago, Monto, Moneda,
            FechaPago, FechaVencimiento, OperacionNumero, EntidadFinanciera,
            ReferenciaExterna, AccionAplicacion, AplicarAlConfirmar, AplicadoSuscripcion,
            FechaAplicacion, UsuarioAplicacion, TipoCobroObjetivo, PlanComercialObjetivo,
            TipoPlanObjetivo, SedesPermitidasObjetivo, EspaciosPermitidosObjetivo, UsuariosPermitidosObjetivo, FechaInicioPlanObjetivo,
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
            NULL, NULL, NULLIF(@TipoCobroObjetivoNorm, N''), NULLIF(@PlanComercialObjetivoNorm, N''),
            @TipoPlanObjetivo, @SedesPermitidasObjetivo, @EspaciosPermitidosObjetivo, @UsuariosPermitidosObjetivo, @FechaInicioPlanObjetivo,
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

            UPDATE dbo.NegociosSuscripcion
            SET PlanComercial = @PlanComercialObjetivoNorm,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE NegocioId = @NegocioId;

            UPDATE dbo.Negocios
            SET TipoPlan = @TipoPlanObjetivo,
                SedesPermitidas = @SedesPermitidasObjetivo,
                EspaciosPermitidos = @EspaciosPermitidosObjetivo,
                UsuariosPermitidos = @UsuariosPermitidosObjetivo
            WHERE Id = @NegocioId;

            SELECT TOP (1)
                @MovimientoAplicadoId = m.Id
            FROM dbo.NegociosSuscripcionMovimiento m
            WHERE m.NegocioId = @NegocioId
              AND m.TipoMovimiento = @AccionAplicacionNorm
            ORDER BY m.FechaCreacion DESC, m.Id DESC;

            UPDATE dbo.NegociosSuscripcionMovimiento
            SET PlanComercialAnterior = @PlanComercialAnterior,
                PlanComercialNuevo = @PlanComercialObjetivoNorm,
                TipoPlanAnterior = @TipoPlanAnterior,
                TipoPlanNuevo = @TipoPlanObjetivo,
                SedesPermitidasAnterior = @SedesPermitidasAnterior,
                SedesPermitidasNuevo = @SedesPermitidasObjetivo,
                EspaciosPermitidosAnterior = @EspaciosPermitidosAnterior,
                EspaciosPermitidosNuevo = @EspaciosPermitidosObjetivo,
                UsuariosPermitidosAnterior = @UsuariosPermitidosAnterior,
                UsuariosPermitidosNuevo = @UsuariosPermitidosObjetivo
            WHERE Id = @MovimientoAplicadoId;

            UPDATE dbo.NegociosSuscripcionPago
            SET AplicadoSuscripcion = 1,
                NegocioSuscripcionMovimientoId = @MovimientoAplicadoId,
                FechaAplicacion = SYSUTCDATETIME(),
                UsuarioAplicacion = COALESCE(@Usuario, N'sistema'),
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE Id = @PagoId;
        END

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
