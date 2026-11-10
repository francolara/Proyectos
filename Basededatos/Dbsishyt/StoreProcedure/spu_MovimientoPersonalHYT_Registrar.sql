-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_MovimientoPersonalHYT_Registrar para pagos del personal.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Permite omitir la fila de respuesta al registrar movimientos desde el cierre de planilla.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MovimientoPersonalHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @TrabajadorId BIGINT,
    @Concepto VARCHAR(24),
    @Fecha DATE,
    @Descripcion NVARCHAR(300),
    @Importe DECIMAL(20,4),
    @MonedaCodigo NVARCHAR(3),
    @Frecuencia VARCHAR(10) = NULL,
    @PeriodoDesde DATE = NULL,
    @PeriodoHasta DATE = NULL,
    @Referencia NVARCHAR(100) = NULL,
    @CeseId BIGINT = NULL,
    @DevolverId BIT = 1
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar movimientos del personal.', 1;

        SET @Descripcion = NULLIF(TRIM(@Descripcion), N'');
        SET @Referencia = NULLIF(TRIM(@Referencia), N'');
        IF @OperacionId IS NULL OR @Fecha IS NULL OR @Descripcion IS NULL OR @Importe IS NULL
           OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2) OR @Concepto IS NULL
           OR @Concepto NOT IN ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION',
               'VACACIONES_PENDIENTES', 'ADELANTO', 'ADELANTO_GRATIFICACION', 'DESCUENTO', 'PAGO')
            THROW 51000, N'Revise el concepto, la fecha y el importe.', 1;

        IF (@Frecuencia IS NULL AND (@PeriodoDesde IS NOT NULL OR @PeriodoHasta IS NOT NULL))
           OR (@Frecuencia IS NOT NULL AND (@Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
               OR @PeriodoDesde IS NULL OR @PeriodoHasta IS NULL OR @PeriodoHasta < @PeriodoDesde))
            THROW 51000, N'Revise la frecuencia y las fechas de cobertura.', 1;

        IF @Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION', 'PAGO') AND @Referencia IS NULL
            THROW 51000, N'Indique la referencia del pago o adelanto.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        DECLARE @MovimientoId BIGINT;
        SELECT @MovimientoId = movimiento.MovimientoId
        FROM dbo.MovimientoPersonalHYT AS movimiento
        WHERE movimiento.OperacionId = @OperacionId;
        IF @MovimientoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.MovimientoPersonalHYT AS movimiento
                WHERE movimiento.MovimientoId = @MovimientoId AND movimiento.UsuarioId = @UsuarioId
                  AND movimiento.TrabajadorId = @TrabajadorId AND movimiento.Concepto = @Concepto
                  AND movimiento.Fecha = @Fecha AND movimiento.Descripcion = @Descripcion
                  AND movimiento.Importe = @Importe AND movimiento.MonedaCodigo = @MonedaCodigo
                  AND movimiento.Frecuencia IS NOT DISTINCT FROM @Frecuencia
                  AND movimiento.PeriodoDesde IS NOT DISTINCT FROM @PeriodoDesde
                  AND movimiento.PeriodoHasta IS NOT DISTINCT FROM @PeriodoHasta
                  AND movimiento.Referencia IS NOT DISTINCT FROM @Referencia
                  AND movimiento.CeseId IS NOT DISTINCT FROM @CeseId)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Mes DATE = DATEFROMPARTS(YEAR(@Fecha), MONTH(@Fecha), 1);
            IF EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes >= @Mes)
                THROW 51000, N'El movimiento corresponde a un mes cerrado o anterior al ultimo cierre.', 1;

            DECLARE @Tipo VARCHAR(10), @Ingreso DATE;
            SELECT @Tipo = trabajador.Tipo, @Ingreso = trabajador.FechaIngreso
            FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
            WHERE trabajador.TrabajadorId = @TrabajadorId;
            IF @Tipo IS NULL OR @Fecha < @Ingreso
                THROW 51000, N'El trabajador o la fecha del movimiento no son validos.', 1;
            IF (@Tipo = 'EXTERNO' AND @Concepto IN ('REMUNERACION', 'GRATIFICACION',
                'ADELANTO_GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES'))
                OR (@Tipo = 'PLANILLA' AND @Concepto = 'SERVICIO')
                THROW 51000, N'El concepto no corresponde al tipo de trabajador.', 1;

            IF (@Concepto IN ('LIQUIDACION', 'VACACIONES_PENDIENTES') AND @CeseId IS NULL)
                OR (@CeseId IS NOT NULL AND NOT EXISTS (
                    SELECT 1 FROM dbo.CeseTrabajadorHYT AS cese
                    WHERE cese.CeseId = @CeseId AND cese.TrabajadorId = @TrabajadorId AND cese.FechaCese <= @Fecha))
                THROW 51000, N'Seleccione un cese valido del trabajador.', 1;

            IF @Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION', 'PAGO', 'DESCUENTO')
                AND EXISTS (
                    SELECT 1 FROM dbo.MovimientoPersonalHYT AS posterior
                    WHERE posterior.TrabajadorId = @TrabajadorId AND posterior.MonedaCodigo = @MonedaCodigo
                      AND posterior.Fecha > @Fecha AND posterior.Concepto IN ('PAGO', 'DESCUENTO'))
                THROW 51000, N'Hay pagos o descuentos posteriores a la fecha del movimiento.', 1;

            IF @Concepto IN ('PAGO', 'DESCUENTO')
            BEGIN
                DECLARE @Saldo DECIMAL(38,2);
                SELECT @Saldo = COALESCE(SUM(CASE WHEN movimiento.Concepto IN
                    ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                    THEN movimiento.Importe ELSE -movimiento.Importe END), 0)
                FROM dbo.MovimientoPersonalHYT AS movimiento
                WHERE movimiento.TrabajadorId = @TrabajadorId AND movimiento.MonedaCodigo = @MonedaCodigo
                  AND movimiento.Fecha <= @Fecha;
                IF @Importe > @Saldo
                    THROW 51000, N'El importe excede el saldo del trabajador a esa fecha.', 1;
            END;

            INSERT INTO dbo.MovimientoPersonalHYT
                (TrabajadorId, CeseId, TipoTrabajador, Concepto, Fecha, Descripcion, Importe,
                 MonedaCodigo, Frecuencia, PeriodoDesde, PeriodoHasta, Referencia, OperacionId, UsuarioId)
            VALUES (@TrabajadorId, @CeseId, @Tipo, @Concepto, @Fecha, @Descripcion, @Importe,
                    @MonedaCodigo, @Frecuencia, @PeriodoDesde, @PeriodoHasta, @Referencia, @OperacionId, @UsuarioId);

            SET @MovimientoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @MovimientoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'MovimientoPersonalHYT',
                 @RegistroId, N'REGISTRAR', @EstadoNuevo = @Concepto;
        END;

        COMMIT TRANSACTION;
        IF @DevolverId = 1 SELECT @MovimientoId AS MovimientoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
