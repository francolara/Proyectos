-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta costos, saldos y datos de pago con filtros de fecha, estado y busqueda.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CostoProveedor_Listar
    @UsuarioId NVARCHAR(128),
    @PresupuestoId BIGINT = NULL,
    @ProveedorId BIGINT = NULL,
    @SoloPendientes BIT = 0,
    @Pagina INT = 1,
    @CostoProveedorId BIGINT = NULL,
    @Desde DATE = NULL,
    @Hasta DATE = NULL,
    @EstadoFiltro VARCHAR(10) = NULL,
    @Buscar NVARCHAR(180) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR') = 0
           AND (@PresupuestoId IS NULL
                OR dbo.fn_HYTPermiso(@UsuarioId, N'COSTOS_PROVEEDOR.CONSULTAR') = 0
                OR dbo.fn_HYTPermiso(@UsuarioId, N'PRESUPUESTOS.CONSULTAR') = 0)
            THROW 51000, N'No tiene permiso para consultar costos de proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;
        IF @Desde > @Hasta OR @Hasta = '99991231'
           OR (@EstadoFiltro IS NOT NULL AND @EstadoFiltro NOT IN ('PENDIENTES', 'PAGADOS', 'TODOS'))
            THROW 51000, N'Revise el rango de fechas y el estado.', 1;

        DECLARE @DesdeUtc DATETIME2(0) = DATEADD(HOUR, 5, CONVERT(DATETIME2(0), @Desde));
        DECLARE @HastaUtc DATETIME2(0) = DATEADD(HOUR, 5, DATEADD(DAY, 1, CONVERT(DATETIME2(0), @Hasta)));
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA');
        DECLARE @VerPago BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR');

        SELECT costo.CostoProveedorId, costo.PresupuestoId,
               COALESCE(presupuesto.NumeroPresupuesto, presupuesto.NumeroInterno) AS NumeroInterno,
               CONVERT(DATE, DATEADD(HOUR, -5, costo.FechaRegistro)) AS FechaCosto,
               costo.ProveedorId, proveedor.RazonSocial, costo.Descripcion, costo.MonedaCodigo,
               CASE WHEN @CostoProveedorId IS NOT NULL AND @VerPago = 1
                    THEN proveedor.Banco END AS Banco,
               CASE WHEN @CostoProveedorId IS NOT NULL AND @VerPago = 1 AND @VerCuenta = 1
                    THEN proveedor.NumeroCuenta END AS NumeroCuenta,
               costo.DocumentoSustento, costo.Importe, saldo.Pagado,
               costo.Importe - saldo.Pagado AS Pendiente,
               CASE WHEN saldo.Pagado = 0 THEN N'PENDIENTE'
                    WHEN saldo.Pagado = costo.Importe THEN N'PAGADO' ELSE N'PARCIAL' END AS Estado,
               COUNT_BIG(1) OVER () AS Total
        FROM dbo.CostoProveedorPresupuesto AS costo
        INNER JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId = costo.PresupuestoId
        INNER JOIN dbo.ProveedorHYT AS proveedor ON proveedor.ProveedorId = costo.ProveedorId
        OUTER APPLY (
            SELECT CONVERT(DECIMAL(18,2), COALESCE(SUM(pago.Importe), 0)) AS Pagado
            FROM dbo.PagoProveedorHYT AS pago
            WHERE pago.CostoProveedorId = costo.CostoProveedorId
        ) AS saldo
        WHERE (@PresupuestoId IS NULL OR costo.PresupuestoId = @PresupuestoId)
          AND (@ProveedorId IS NULL OR costo.ProveedorId = @ProveedorId)
          AND (@CostoProveedorId IS NULL OR costo.CostoProveedorId = @CostoProveedorId)
          AND (@DesdeUtc IS NULL OR costo.FechaRegistro >= @DesdeUtc)
          AND (@HastaUtc IS NULL OR costo.FechaRegistro < @HastaUtc)
          AND (@Buscar IS NULL
               OR CHARINDEX(@Buscar, proveedor.RazonSocial COLLATE Latin1_General_100_CI_AI) > 0
               OR CHARINDEX(@Buscar, proveedor.Ruc) > 0
               OR CHARINDEX(@Buscar, costo.Descripcion COLLATE Latin1_General_100_CI_AI) > 0
               OR CHARINDEX(@Buscar, presupuesto.NumeroPresupuesto) > 0
               OR CHARINDEX(@Buscar, presupuesto.NumeroInterno) > 0)
          AND (@EstadoFiltro = 'TODOS'
               OR (@EstadoFiltro = 'PAGADOS' AND saldo.Pagado = costo.Importe)
               OR (@EstadoFiltro = 'PENDIENTES' AND saldo.Pagado < costo.Importe)
               OR (@EstadoFiltro IS NULL AND (@SoloPendientes = 0 OR saldo.Pagado < costo.Importe)))
        ORDER BY costo.CostoProveedorId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
