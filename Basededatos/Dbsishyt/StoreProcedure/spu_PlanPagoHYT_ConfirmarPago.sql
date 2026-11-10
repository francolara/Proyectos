-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Confirma la planificacion y registra sus movimientos de forma atomica.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_ConfirmarPago
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE,
    @VersionFila BINARY(8),
    @FechaPago DATE,
    @ReferenciaPago NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
           OR dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.EDITAR') = 0
            THROW 51000, N'No tiene permiso para confirmar el pago de planilla.', 1;

        SET @ReferenciaPago = NULLIF(TRIM(@ReferenciaPago), N'');
        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @Desde IS NULL OR @VersionFila IS NULL OR @FechaPago IS NULL
           OR @FechaPago < @Desde OR @FechaPago > @Hoy OR @ReferenciaPago IS NULL
            THROW 51000, N'Revise la fecha y el numero de operacion.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        DECLARE @PlanId BIGINT, @Actual BINARY(8), @Hasta DATE, @FechaActual DATE, @ReferenciaActual NVARCHAR(100);
        SELECT @PlanId = periodo.PlanId, @Actual = periodo.VersionFila, @Hasta = periodo.Hasta,
               @FechaActual = periodo.FechaPago, @ReferenciaActual = periodo.ReferenciaPago
        FROM dbo.PlanPagoHYT AS periodo WITH (UPDLOCK, HOLDLOCK)
        WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;

        IF @PlanId IS NULL
            THROW 51000, N'Guarde la planificacion antes de confirmar el pago.', 1;
        IF @FechaActual IS NOT NULL
        BEGIN
            IF @FechaActual <> @FechaPago OR @ReferenciaActual <> @ReferenciaPago
                THROW 51000, N'La planificacion ya esta pagada y cerrada.', 1;
            COMMIT TRANSACTION;
            SELECT @PlanId AS PlanId;
            RETURN;
        END;
        IF @Actual <> @VersionFila
            THROW 51000, N'La planificacion cambio. Vuelva a abrirla y revise los importes.', 1;

        IF EXISTS (
            SELECT 1 FROM dbo.PlanPagoDetalleHYT AS detalle
            INNER JOIN dbo.MovimientoPersonalHYT AS movimiento ON movimiento.TrabajadorId = detalle.TrabajadorId
            WHERE detalle.PlanId = @PlanId AND movimiento.MonedaCodigo = N'PEN'
              AND movimiento.Concepto IN ('REMUNERACION', 'SERVICIO', 'DESCUENTO', 'PAGO')
              AND movimiento.PeriodoDesde <= @Hasta AND movimiento.PeriodoHasta >= @Desde)
            THROW 51000, N'Hay movimientos de personal que ya cubren este periodo. Revise el estado de cuenta.', 1;

        DECLARE @TrabajadorId BIGINT, @Tipo VARCHAR(10), @Bruto DECIMAL(20,4),
                @Descuento DECIMAL(20,4), @Neto DECIMAL(20,4), @Concepto VARCHAR(24),
                @Operacion UNIQUEIDENTIFIER, @Descripcion NVARCHAR(300) = CONCAT(N'Planificacion de Planilla #', @PlanId);
        DECLARE trabajadores CURSOR LOCAL FAST_FORWARD FOR
            SELECT detalle.TrabajadorId, detalle.Tipo, detalle.Base + detalle.ImporteExtra, detalle.Descuento
            FROM dbo.PlanPagoDetalleHYT AS detalle
            WHERE detalle.PlanId = @PlanId
            ORDER BY detalle.TrabajadorId;

        OPEN trabajadores;
        FETCH NEXT FROM trabajadores INTO @TrabajadorId, @Tipo, @Bruto, @Descuento;
        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @Concepto = CASE WHEN @Tipo = 'PLANILLA' THEN 'REMUNERACION' ELSE 'SERVICIO' END;
            SET @Neto = @Bruto - @Descuento;

            IF @Bruto > 0
            BEGIN
                SET @Operacion = NEWID();
                EXEC dbo.spu_MovimientoPersonalHYT_Registrar @UsuarioId, @Operacion, @TrabajadorId, @Concepto,
                     @FechaPago, @Descripcion, @Bruto, N'PEN', @Frecuencia, @Desde, @Hasta, @DevolverId = 0;
            END;
            IF @Descuento > 0
            BEGIN
                SET @Operacion = NEWID();
                EXEC dbo.spu_MovimientoPersonalHYT_Registrar @UsuarioId, @Operacion, @TrabajadorId, 'DESCUENTO',
                     @FechaPago, @Descripcion, @Descuento, N'PEN', @Frecuencia, @Desde, @Hasta, @DevolverId = 0;
            END;
            IF @Neto > 0
            BEGIN
                SET @Operacion = NEWID();
                EXEC dbo.spu_MovimientoPersonalHYT_Registrar @UsuarioId, @Operacion, @TrabajadorId, 'PAGO',
                     @FechaPago, @Descripcion, @Neto, N'PEN', @Frecuencia, @Desde, @Hasta, @ReferenciaPago, @DevolverId = 0;
            END;

            FETCH NEXT FROM trabajadores INTO @TrabajadorId, @Tipo, @Bruto, @Descuento;
        END;
        CLOSE trabajadores;
        DEALLOCATE trabajadores;

        UPDATE dbo.PlanPagoHYT SET FechaPago = @FechaPago, ReferenciaPago = @ReferenciaPago WHERE PlanId = @PlanId;
        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PlanId);
        SET @Operacion = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (SELECT @ReferenciaPago AS ReferenciaPago FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @Operacion, N'PlanPagoHYT', @RegistroId,
             N'CONFIRMAR_PAGO', @EstadoNuevo = N'PAGADA', @FechaEvento = @FechaPago, @CambiosJson = @Cambios;

        COMMIT TRANSACTION;
        SELECT @PlanId AS PlanId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
