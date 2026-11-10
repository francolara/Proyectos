-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta el historial de abonos con frecuencia y periodo de cobertura.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PagoProveedor_Listar
    @UsuarioId NVARCHAR(128),
    @CostoProveedorId BIGINT,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar pagos a proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        SELECT pago.PagoProveedorId, pago.Importe, pago.FechaPago, pago.Referencia,
               pago.Frecuencia, pago.PeriodoDesde, pago.PeriodoHasta,
               pago.FechaRegistro, pago.UsuarioRegistro, costo.MonedaCodigo,
               COUNT_BIG(1) OVER () AS Total
        FROM dbo.PagoProveedorHYT AS pago
        INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId = pago.CostoProveedorId
        WHERE pago.CostoProveedorId = @CostoProveedorId
        ORDER BY pago.FechaPago DESC, pago.PagoProveedorId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
