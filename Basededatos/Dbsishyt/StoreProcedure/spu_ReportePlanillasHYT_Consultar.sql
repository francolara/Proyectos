-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Concilia el mes de personal y proveedores por moneda.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ReportePlanillasHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE,
    @ProveedorId BIGINT = NULL,
    @PresupuestoId BIGINT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'REPORTES_PLANILLAS.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar reportes de planillas.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hasta DATE = EOMONTH(@Mes);
        DECLARE @HastaUtc DATETIME2(0) = DATEADD(HOUR, 5, DATEADD(DAY, 1, CONVERT(DATETIME2(0), @Hasta)));
        DECLARE @VerPersonal BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR');
        DECLARE @VerProveedores BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR');

        SELECT (
            SELECT @Mes AS Mes, @VerPersonal AS VerPersonal, @VerProveedores AS VerProveedores,
                   CONVERT(BIT, CASE WHEN @VerPersonal = 1 AND EXISTS (
                       SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes = @Mes)
                       THEN 1 ELSE 0 END) AS Cerrado,
                   JSON_QUERY((
                       SELECT movimiento.TrabajadorId, trabajador.NombreCompleto,
                              movimiento.TipoTrabajador, movimiento.MonedaCodigo,
                              SUM(CASE WHEN movimiento.Fecha < @Mes THEN signo.Neto ELSE 0 END) AS SaldoAnterior,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND signo.Neto > 0 THEN movimiento.Importe ELSE 0 END) AS Devengado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION')
                                  THEN movimiento.Importe ELSE 0 END) AS Adelantado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'DESCUENTO' THEN movimiento.Importe ELSE 0 END) AS Descontado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'PAGO' THEN movimiento.Importe ELSE 0 END) AS Pagado,
                              SUM(signo.Neto) AS Saldo
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = movimiento.TrabajadorId
                       CROSS APPLY (VALUES (CASE WHEN movimiento.Concepto IN ('REMUNERACION', 'SERVICIO',
                           'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                           THEN movimiento.Importe ELSE -movimiento.Importe END)) AS signo (Neto)
                       WHERE @VerPersonal = 1 AND movimiento.Fecha <= @Hasta
                       GROUP BY movimiento.TrabajadorId, trabajador.NombreCompleto, movimiento.TipoTrabajador, movimiento.MonedaCodigo
                       HAVING MAX(movimiento.Fecha) >= @Mes OR SUM(signo.Neto) <> 0
                       ORDER BY movimiento.TipoTrabajador, trabajador.NombreCompleto, movimiento.TrabajadorId, movimiento.MonedaCodigo
                       FOR JSON PATH
                   )) AS Personal,
                   JSON_QUERY((
                       SELECT costo.ProveedorId, proveedor.RazonSocial, costo.PresupuestoId, presupuesto.NumeroInterno,
                              costo.MonedaCodigo, SUM(costo.Importe) AS Costo,
                              SUM(COALESCE(abonos.PagadoMes, 0)) AS PagadoMes,
                              SUM(COALESCE(abonos.PagadoAcumulado, 0)) AS PagadoAcumulado,
                              SUM(CONVERT(DECIMAL(19,2), costo.Importe - COALESCE(abonos.PagadoAcumulado, 0))) AS Pendiente
                       FROM dbo.CostoProveedorPresupuesto AS costo
                       INNER JOIN dbo.ProveedorHYT AS proveedor ON proveedor.ProveedorId = costo.ProveedorId
                       INNER JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId = costo.PresupuestoId
                       OUTER APPLY (
                           SELECT CONVERT(DECIMAL(18,2), SUM(CASE WHEN pago.FechaPago >= @Mes THEN pago.Importe ELSE 0 END)) AS PagadoMes,
                                  CONVERT(DECIMAL(18,2), SUM(pago.Importe)) AS PagadoAcumulado
                           FROM dbo.PagoProveedorHYT AS pago
                           WHERE pago.CostoProveedorId = costo.CostoProveedorId AND pago.FechaPago <= @Hasta
                       ) AS abonos
                       WHERE @VerProveedores = 1
                         AND (costo.FechaRegistro < @HastaUtc OR abonos.PagadoAcumulado IS NOT NULL)
                         AND (@ProveedorId IS NULL OR costo.ProveedorId = @ProveedorId)
                         AND (@PresupuestoId IS NULL OR costo.PresupuestoId = @PresupuestoId)
                       GROUP BY costo.ProveedorId, proveedor.RazonSocial, costo.PresupuestoId, presupuesto.NumeroInterno, costo.MonedaCodigo
                       HAVING SUM(COALESCE(abonos.PagadoMes, 0)) > 0 OR SUM(costo.Importe) > SUM(COALESCE(abonos.PagadoAcumulado, 0))
                       ORDER BY proveedor.RazonSocial, presupuesto.NumeroInterno, costo.MonedaCodigo
                       FOR JSON PATH
                   )) AS Proveedores
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
