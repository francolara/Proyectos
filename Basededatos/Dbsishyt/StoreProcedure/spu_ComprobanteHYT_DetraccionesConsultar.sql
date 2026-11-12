-- Firma: FRANCO LARA | Create date: 10/11/2026 | Listado de detracciones por estado, proveedor y comprobante, con fecha y operación del último pago.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_DetraccionesConsultar
    @UsuarioId NVARCHAR(128),@Estado NVARCHAR(20)=N'PENDIENTE',@Proveedor NVARCHAR(150)=NULL,
    @Busqueda NVARCHAR(150)=NULL,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PAGOS_PROVEEDOR.CONSULTAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar detracciones.',1;
        IF @Estado IS NULL OR @Estado NOT IN(N'TODOS',N'PENDIENTE',N'PAGADA') OR @Pagina NOT BETWEEN 1 AND 100000
            THROW 51601,N'[HYT_VALIDACION] Revisa los filtros del listado.',1;
        SELECT COALESCE((
            SELECT c.ComprobanteId,c.Serie,c.Numero,c.Fecha,c.NombreTercero,c.DocumentoTercero,
                c.ImporteDetraccion,c.MonedaDetraccion,c.CodigoDetraccion,
                COALESCE(pago.Importe,0) AS Pagado,c.ImporteDetraccion-COALESCE(pago.Importe,0) AS Saldo,
                ultimo.Fecha AS FechaUltimoPago,ultimo.Referencia AS UltimaOperacion,
                COUNT_BIG(1) OVER() AS TotalFilas
            FROM dbo.ComprobanteHYT AS c
            OUTER APPLY(SELECT SUM(p.Importe) AS Importe FROM dbo.ComprobanteHYTPago AS p WHERE p.ComprobanteId=c.ComprobanteId AND p.EsDetraccion=1) AS pago
            OUTER APPLY(SELECT TOP (1) pagoReciente.Fecha,pagoReciente.Referencia
                FROM dbo.ComprobanteHYTPago AS pagoReciente WHERE pagoReciente.ComprobanteId=c.ComprobanteId AND pagoReciente.EsDetraccion=1
                ORDER BY pagoReciente.Fecha DESC,pagoReciente.PagoId DESC) AS ultimo
            WHERE c.EsVenta=0 AND c.Tipo<>'07' AND c.ImporteDetraccion>0
                AND (@Estado=N'TODOS' OR (@Estado=N'PAGADA' AND COALESCE(pago.Importe,0)>=c.ImporteDetraccion)
                    OR (@Estado=N'PENDIENTE' AND COALESCE(pago.Importe,0)<c.ImporteDetraccion))
                AND (@Proveedor IS NULL OR CHARINDEX(@Proveedor,c.NombreTercero)>0 OR CHARINDEX(@Proveedor,c.DocumentoTercero)>0)
                AND (@Busqueda IS NULL OR CHARINDEX(@Busqueda,CONCAT(c.Serie,N'-',c.Numero))>0 OR CHARINDEX(@Busqueda,c.CodigoDetraccion)>0
                    OR EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPago AS pagoBusqueda WHERE pagoBusqueda.ComprobanteId=c.ComprobanteId
                        AND pagoBusqueda.EsDetraccion=1 AND CHARINDEX(@Busqueda,pagoBusqueda.Referencia)>0))
            ORDER BY c.Fecha,c.ComprobanteId OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY
            FOR JSON PATH),N'[]');
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_DetraccionesConsultar TO hyt_app_executor;
GO
