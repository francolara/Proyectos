-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta ajustes y datos bancarios autorizados del reporte.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Incluye la confirmacion de pago en la consulta y reporte.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE,
    @Reporte BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar la planificacion.', 1;

        IF @Frecuencia IS NULL OR @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
           OR @Desde IS NULL OR @Desde > '99991224'
           OR (@Frecuencia = 'SEMANAL' AND DATEDIFF(DAY, '19000101', @Desde) % 7 <> 0)
           OR (@Frecuencia = 'QUINCENAL' AND DAY(@Desde) NOT IN (1,16))
           OR (@Frecuencia = 'MENSUAL' AND DAY(@Desde) <> 1)
            THROW 51000, N'Seleccione el inicio valido del periodo.', 1;

        DECLARE @Hasta DATE = CASE WHEN @Frecuencia = 'SEMANAL' THEN DATEADD(DAY, 6, @Desde)
            WHEN @Frecuencia = 'QUINCENAL' AND DAY(@Desde) = 1 THEN DATEADD(DAY, 14, @Desde)
            ELSE EOMONTH(@Desde) END;

        DECLARE @PlanId BIGINT, @Version BINARY(8), @FechaPago DATE, @ReferenciaPago NVARCHAR(100);
        SELECT @PlanId = periodo.PlanId, @Version = periodo.VersionFila,
               @FechaPago = periodo.FechaPago, @ReferenciaPago = periodo.ReferenciaPago
        FROM dbo.PlanPagoHYT AS periodo WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;

        DECLARE @VerCuenta BIT = CASE WHEN @Reporte = 1
            AND dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 1 THEN 1 ELSE 0 END;
        DECLARE @Filas NVARCHAR(MAX);
        IF @PlanId IS NOT NULL
            SET @Filas = (
                SELECT detalle.TrabajadorId, detalle.Nombre, detalle.Tipo, detalle.Base,
                       detalle.MinutosExtra, detalle.ImporteExtra, detalle.Descuento,
                       JSON_QUERY(detalle.Ajustes) AS Ajustes,
                       CASE WHEN @VerCuenta = 1 THEN trabajador.Banco END AS Banco,
                       CASE WHEN @VerCuenta = 1 THEN trabajador.NumeroCuenta END AS Cuenta,
                       CASE WHEN @VerCuenta = 1 THEN trabajador.CuentaInterbancaria END AS Cci
                FROM dbo.PlanPagoDetalleHYT AS detalle
                INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = detalle.TrabajadorId
                WHERE detalle.PlanId = @PlanId
                ORDER BY detalle.Tipo, detalle.Nombre, detalle.TrabajadorId FOR JSON PATH);
        ELSE
            SET @Filas = (
                SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre, trabajador.Tipo,
                       COALESCE(trabajador.SueldoMensual, 0) AS Base, 0 AS MinutosExtra,
                       0 AS ImporteExtra, 0 AS Descuento
                FROM dbo.TrabajadorHYT AS trabajador
                WHERE trabajador.Activo = 1 AND trabajador.FrecuenciaPago = @Frecuencia AND trabajador.FechaIngreso <= @Hasta
                ORDER BY trabajador.Tipo, trabajador.NombreCompleto, trabajador.TrabajadorId FOR JSON PATH);

        SELECT (SELECT @PlanId AS PlanId, @Version AS VersionFila, @Frecuencia AS Frecuencia,
            @Desde AS Desde, @Hasta AS Hasta, @VerCuenta AS VerCuenta,
            @FechaPago AS FechaPago, @ReferenciaPago AS ReferenciaPago, JSON_QUERY(COALESCE(@Filas, N'[]')) AS Filas
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
