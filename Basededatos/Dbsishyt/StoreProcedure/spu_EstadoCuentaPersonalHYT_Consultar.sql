-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_EstadoCuentaPersonalHYT_Consultar para Planillas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Incluye fecha y motivo de reingreso en el historial.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_EstadoCuentaPersonalHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar pagos del personal.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hasta DATE = EOMONTH(@Mes);
        SELECT (
            SELECT trabajador.TrabajadorId, trabajador.NombreCompleto, trabajador.Tipo, trabajador.Activo,
                   @Mes AS Mes, CONVERT(BIT, CASE WHEN cierre.Mes IS NULL THEN 0 ELSE 1 END) AS Cerrado,
                   JSON_QUERY((
                       SELECT movimiento.MonedaCodigo,
                              SUM(CASE WHEN movimiento.Fecha < @Mes THEN signo.Neto ELSE 0 END) AS SaldoAnterior,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND signo.Neto > 0 THEN movimiento.Importe ELSE 0 END) AS Devengado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION') THEN movimiento.Importe ELSE 0 END) AS Adelantado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'DESCUENTO' THEN movimiento.Importe ELSE 0 END) AS Descontado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'PAGO' THEN movimiento.Importe ELSE 0 END) AS Pagado,
                              SUM(signo.Neto) AS Saldo
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       CROSS APPLY (VALUES (CASE WHEN movimiento.Concepto IN ('REMUNERACION', 'SERVICIO',
                           'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                           THEN movimiento.Importe ELSE -movimiento.Importe END)) AS signo (Neto)
                       WHERE movimiento.TrabajadorId = trabajador.TrabajadorId AND movimiento.Fecha <= @Hasta
                       GROUP BY movimiento.MonedaCodigo
                       ORDER BY movimiento.MonedaCodigo
                       FOR JSON PATH
                   )) AS Saldos,
                   JSON_QUERY((
                       SELECT movimiento.MovimientoId, movimiento.TipoTrabajador, movimiento.Concepto,
                              movimiento.Fecha, movimiento.Descripcion, movimiento.Importe, movimiento.MonedaCodigo,
                              movimiento.Frecuencia, movimiento.PeriodoDesde, movimiento.PeriodoHasta,
                              movimiento.Referencia, movimiento.CeseId
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       WHERE movimiento.TrabajadorId = trabajador.TrabajadorId AND movimiento.Fecha BETWEEN @Mes AND @Hasta
                       ORDER BY movimiento.Fecha DESC, movimiento.MovimientoId DESC
                       FOR JSON PATH
                   )) AS Movimientos,
                   JSON_QUERY((
                       SELECT cese.CeseId, cese.FechaCese, cese.UltimoDiaTrabajado, cese.Motivo, cese.FechaReingreso, cese.MotivoReingreso
                       FROM dbo.CeseTrabajadorHYT AS cese
                       WHERE cese.TrabajadorId = trabajador.TrabajadorId
                       ORDER BY cese.FechaCese DESC, cese.CeseId DESC
                       FOR JSON PATH
                   )) AS Ceses
            FROM dbo.TrabajadorHYT AS trabajador
            LEFT JOIN dbo.CierrePersonalHYT AS cierre ON cierre.Mes = @Mes
            WHERE trabajador.TrabajadorId = @TrabajadorId
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
