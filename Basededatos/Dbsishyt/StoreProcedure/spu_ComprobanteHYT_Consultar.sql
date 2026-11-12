-- Firma: FRANCO LARA | Create date: 10/11/2026 | Operaciones de provisiones de compras y comprobantes de venta. Devuelve arreglos vacíos para consultas sin coincidencias y lista clientes y proveedores activos desde la apertura de la ayuda y consulta el IGV vigente y la fecha de vencimiento. Identifica facturas sin factoring y consulta el catálogo de detracciones.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Filtra ventas con o sin factoring y devuelve la fecha de desembolso registrada en el lote y expone la detracción de ventas y su consulta por material, la orden vinculada y su confirmación vigente. Busca trabajos disponibles y presupuestos de venta por número o descripción y ofrece las facturas del cliente y los catálogos SUNAT.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_Consultar
    @UsuarioId NVARCHAR(128), @EsVenta BIT, @Id BIGINT=NULL, @Desde DATE=NULL, @Hasta DATE=NULL,
    @Tipo CHAR(2)=NULL, @Busqueda NVARCHAR(150)=NULL, @Pagina INT=1, @ClienteId BIGINT=NULL, @PresupuestoId BIGINT=NULL, @SinFactoring BIT=0, @ConFactoring BIT=NULL,
    @DetalleMateriales NVARCHAR(MAX)=NULL, @BuscarOrdenes BIT=0, @BuscarFacturas BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @EsVenta=1 THEN N'FACTURADO.CONSULTAR' ELSE N'PAGOS_PROVEEDOR.CONSULTAR' END)=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @Pagina NOT BETWEEN 1 AND 100000 OR @Desde>@Hasta
            THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas y la página.',1;
        SELECT (
            SELECT
                (SELECT Ruc FROM dbo.ConfiguracionHYT WHERE Id=1) AS RucHYT,
                (SELECT TOP (1) tasa.Valor FROM dbo.TasaIGV AS tasa
                    WHERE tasa.FechaDesde<=@Desde AND (tasa.FechaHasta IS NULL OR tasa.FechaHasta>@Desde)
                    ORDER BY tasa.FechaDesde DESC) AS TasaIgv,
                JSON_QUERY(COALESCE(CASE WHEN @Id IS NULL AND @Desde IS NOT NULL AND @Hasta IS NULL THEN (
                    SELECT catalogo.Codigo,catalogo.Nombre,catalogo.Porcentaje
                    FROM dbo.DetraccionCatalogoHYT AS catalogo WHERE catalogo.Activo=1 ORDER BY catalogo.Codigo FOR JSON PATH)
                END,N'[]')) AS CatalogoDetracciones,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=1 AND @Id IS NULL AND @Desde IS NOT NULL AND @Hasta IS NULL THEN (
                    SELECT motivo.Codigo,motivo.Descripcion FROM dbo.MotivoNotaCreditoHYT AS motivo
                    ORDER BY motivo.Codigo FOR JSON PATH)
                END,N'[]')) AS MotivosNotaCredito,
                JSON_QUERY(COALESCE(CASE WHEN @Id IS NULL AND @Desde IS NOT NULL AND @Hasta IS NULL THEN (
                    SELECT unidad.Codigo,unidad.Descripcion,unidad.CodigoSunat FROM dbo.UnidadMedida AS unidad
                    WHERE unidad.Activo=1 ORDER BY unidad.Codigo FOR JSON PATH)
                END,N'[]')) AS OpcionesUnidad,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=1 AND @BuscarFacturas=1 AND @ClienteId IS NOT NULL THEN (
                    SELECT TOP (20) CONCAT(factura.Serie,N'-',factura.Numero) AS Referencia,factura.Fecha,factura.Moneda,factura.Importe
                    FROM dbo.ComprobanteHYT AS factura
                    WHERE factura.EsVenta=1 AND factura.Tipo='01' AND factura.ClienteId=@ClienteId
                        AND factura.Serie IS NOT NULL AND factura.Numero IS NOT NULL
                        AND factura.EstadoSunat NOT IN(N'RECHAZADO',N'ANULADO',N'BAJA')
                        AND factura.Fecha BETWEEN @Desde AND @Hasta
                        AND (NULLIF(@Busqueda,N'') IS NULL OR CHARINDEX(@Busqueda,CONCAT(factura.Serie,N'-',factura.Numero))>0)
                    ORDER BY factura.Fecha DESC,factura.ComprobanteId DESC FOR JSON PATH)
                END,N'[]')) AS FacturasNotaCredito,
                JSON_QUERY(CASE WHEN @EsVenta=1 AND ISJSON(@DetalleMateriales,ARRAY)=1 THEN (
                    SELECT material.Codigo,material.Porcentaje FROM dbo.fn_ComprobanteHYT_DetraccionMaterial(@DetalleMateriales) AS material
                    FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) END) AS DetraccionMaterial,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=1 AND @BuscarOrdenes=1 AND @ClienteId IS NOT NULL THEN (
                    SELECT TOP (30) orden.OrdenCompraId,orden.NumeroOCCliente,orden.FechaEmisionOC,p.NumeroPresupuesto,
                        confirmacion.NumeroConfirmacion
                    FROM dbo.OrdenCompraCliente AS orden
                    INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=orden.PresupuestoId AND p.ClienteId=@ClienteId
                    LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=orden.OrdenCompraId AND confirmacion.Vigente=1
                    WHERE orden.Estado<>N'CANCELADA' AND (NULLIF(@Busqueda,N'') IS NULL
                        OR CHARINDEX(@Busqueda,orden.NumeroOCCliente)>0 OR CHARINDEX(@Busqueda,p.NumeroPresupuesto)>0
                        OR CHARINDEX(@Busqueda,confirmacion.NumeroConfirmacion)>0)
                    ORDER BY orden.FechaEmisionOC DESC,orden.OrdenCompraId DESC FOR JSON PATH) END,N'[]')) AS OpcionesOrdenCompra,
                JSON_QUERY(CASE WHEN @Id IS NULL THEN NULL ELSE (
                    SELECT c.ComprobanteId,c.Tipo,c.Serie,c.Numero,c.Fecha,c.FechaVencimiento,c.Moneda,c.RucHYT,c.DocumentoTercero,c.NombreTercero,
                        c.ClienteId,c.ProveedorId,c.OrdenCompraId,orden.NumeroOCCliente,confirmacion.NumeroConfirmacion,c.FormaPago,c.Importe,c.ImporteDetraccion,c.DetraccionDocumento,c.MonedaDetraccion,c.CodigoDetraccion,c.PorcentajeDetraccion,catalogo.Nombre AS NombreDetraccion,
                        c.ReferenciaNota,c.MotivoNota,c.CodigoMotivoNota,c.EstadoSunat,CONVERT(BIT,CASE WHEN c.XmlOriginal IS NULL THEN 0 ELSE 1 END) AS Importado,
                        JSON_QUERY(c.Detalle) AS Detalle,
                        JSON_QUERY(COALESCE((SELECT pago.Fecha,pago.EsDetraccion,pago.Importe,pago.Referencia FROM dbo.ComprobanteHYTPago AS pago WHERE pago.ComprobanteId=c.ComprobanteId ORDER BY pago.Fecha,pago.PagoId FOR JSON PATH),N'[]')) AS Pagos,
                        JSON_QUERY(COALESCE((SELECT p.PresupuestoId,p.NumeroPresupuesto FROM dbo.ComprobanteHYTPresupuesto AS vinculo INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=vinculo.PresupuestoId WHERE vinculo.ComprobanteId=c.ComprobanteId FOR JSON PATH),N'[]')) AS Presupuestos,
                        JSON_QUERY(COALESCE((SELECT costo.CostoProveedorId,p.NumeroPresupuesto,costo.Descripcion,costo.Importe FROM dbo.ComprobanteHYTCosto AS vinculo INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId=vinculo.CostoProveedorId INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=costo.PresupuestoId WHERE vinculo.ComprobanteId=c.ComprobanteId FOR JSON PATH),N'[]')) AS Costos
                    FROM dbo.ComprobanteHYT AS c
                    LEFT JOIN dbo.DetraccionCatalogoHYT AS catalogo ON catalogo.Codigo=c.CodigoDetraccion
                    LEFT JOIN dbo.OrdenCompraCliente AS orden ON orden.OrdenCompraId=c.OrdenCompraId
                    LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=orden.OrdenCompraId AND confirmacion.Vigente=1
                    WHERE c.EsVenta=@EsVenta AND c.ComprobanteId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
                END) AS Documento,
                JSON_QUERY(COALESCE(CASE WHEN @Id IS NOT NULL OR @Desde IS NULL OR @Hasta IS NULL OR @BuscarFacturas=1 THEN N'[]' ELSE (
                    SELECT c.ComprobanteId,c.Tipo,c.Serie,c.Numero,c.Fecha,c.FechaVencimiento,c.Moneda,c.NombreTercero,c.Importe,c.ImporteDetraccion,c.DetraccionDocumento,c.MonedaDetraccion,c.PorcentajeDetraccion,c.EstadoSunat,factura.LoteId,lote.FechaDesembolso AS FactoringFechaDesembolso,
                        orden.NumeroOCCliente,confirmacion.NumeroConfirmacion,
                        COALESCE(pagos.Pagado,0) AS Pagado,COALESCE(pagos.Detraido,0) AS Detraido,COUNT_BIG(1) OVER() AS TotalFilas
                    FROM dbo.ComprobanteHYT AS c
                    LEFT JOIN dbo.FactoringFacturaHYT AS factura ON factura.ComprobanteId=c.ComprobanteId
                    LEFT JOIN dbo.FactoringLoteHYT AS lote ON lote.LoteId=factura.LoteId
                    LEFT JOIN dbo.OrdenCompraCliente AS orden ON orden.OrdenCompraId=c.OrdenCompraId
                    LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=orden.OrdenCompraId AND confirmacion.Vigente=1
                    OUTER APPLY (SELECT SUM(CASE WHEN p.EsDetraccion=0 THEN p.Importe END) AS Pagado,SUM(CASE WHEN p.EsDetraccion=1 THEN p.Importe END) AS Detraido FROM dbo.ComprobanteHYTPago AS p WHERE p.ComprobanteId=c.ComprobanteId) AS pagos
                    WHERE c.EsVenta=@EsVenta AND c.Fecha BETWEEN @Desde AND @Hasta AND (@Tipo IS NULL OR c.Tipo=@Tipo)
                      AND (@ConFactoring IS NULL OR (@EsVenta=1 AND
                           ((@ConFactoring=1 AND factura.LoteId IS NOT NULL) OR (@ConFactoring=0 AND factura.LoteId IS NULL))))
                      AND (@ConFactoring IS NOT NULL OR @SinFactoring=0 OR (c.EsVenta=1 AND c.Tipo='01' AND c.Serie IS NOT NULL AND c.Numero IS NOT NULL AND c.Importe>0 AND factura.ComprobanteId IS NULL))
                      AND (@Busqueda IS NULL OR CHARINDEX(@Busqueda,c.NombreTercero)>0 OR CHARINDEX(@Busqueda,c.DocumentoTercero)>0 OR CHARINDEX(@Busqueda,CONCAT(c.Serie,N'-',c.Numero))>0)
                    ORDER BY c.Fecha DESC,c.ComprobanteId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH)
                END,N'[]')) AS Filas,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=0 OR @ClienteId IS NULL THEN N'[]' ELSE (
                    SELECT TOP (100) p.PresupuestoId,p.NumeroPresupuesto,p.Descripcion,r.CodigoMoneda,r.Total,r.RevisionId
                    FROM dbo.Presupuesto AS p
                    CROSS APPLY (SELECT TOP (1) rev.RevisionId,rev.CodigoMoneda,rev.Total,rev.Estado FROM dbo.PresupuestoRevision AS rev WHERE rev.PresupuestoId=p.PresupuestoId ORDER BY rev.NumeroRevision DESC) AS r
                    WHERE p.ClienteId=@ClienteId AND r.Estado<>N'CANCELADO' AND (@PresupuestoId IS NULL OR p.PresupuestoId=@PresupuestoId)
                      AND (NULLIF(@Busqueda,N'') IS NULL OR CHARINDEX(@Busqueda,p.NumeroPresupuesto)>0
                          OR CHARINDEX(@Busqueda,p.Descripcion)>0)
                    ORDER BY p.PresupuestoId DESC FOR JSON PATH)
                END,N'[]')) AS OpcionesPresupuesto,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=1 OR @Id IS NULL THEN N'[]' ELSE (
                    SELECT TOP (100) costo.CostoProveedorId,p.NumeroPresupuesto,costo.Descripcion,costo.Importe
                    FROM dbo.ComprobanteHYT AS c
                    INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.ProveedorId=c.ProveedorId AND costo.MonedaCodigo=c.Moneda
                    INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=costo.PresupuestoId
                    WHERE c.ComprobanteId=@Id AND c.EsVenta=0
                    AND dbo.fn_HYTPermiso(@UsuarioId,N'COSTOS_PROVEEDOR.CONSULTAR')=1
                    AND (NULLIF(@Busqueda,N'') IS NULL OR CHARINDEX(@Busqueda,p.NumeroPresupuesto)>0
                        OR CHARINDEX(@Busqueda,costo.Descripcion)>0)
                    AND NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYTCosto AS v WHERE v.CostoProveedorId=costo.CostoProveedorId)
                    ORDER BY costo.CostoProveedorId DESC FOR JSON PATH)
                END,N'[]')) AS OpcionesCosto,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=1 OR @Id IS NULL OR @Busqueda IS NULL OR LEN(@Busqueda)<2
                    OR dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.CONSULTAR')=0 THEN N'[]' ELSE (
                    SELECT TOP (30) p.PresupuestoId,p.NumeroPresupuesto,p.Descripcion
                    FROM dbo.Presupuesto AS p WHERE CHARINDEX(@Busqueda,p.NumeroPresupuesto)>0 OR CHARINDEX(@Busqueda,p.Descripcion)>0
                    ORDER BY p.PresupuestoId DESC FOR JSON PATH)
                END,N'[]')) AS OpcionesCostoPresupuesto,
                JSON_QUERY(COALESCE(CASE WHEN @EsVenta=0 OR @Busqueda IS NULL THEN N'[]' ELSE (
                    SELECT TOP (30) cliente.ClienteId,cliente.Codigo,cliente.NumeroIdentificacion AS DocumentoTercero,
                        CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT_WS(N' ',cliente.Nombres,cliente.Apellidos) ELSE cliente.RazonSocial END AS NombreTercero
                    FROM dbo.Cliente AS cliente WHERE cliente.Activo=1 AND NULLIF(cliente.NumeroIdentificacion,N'') IS NOT NULL
                      AND (@Busqueda=N'' OR CHARINDEX(@Busqueda,cliente.Codigo)>0 OR CHARINDEX(@Busqueda,cliente.NumeroIdentificacion)>0
                        OR CHARINDEX(@Busqueda,CONCAT_WS(N' ',cliente.RazonSocial,cliente.Nombres,cliente.Apellidos))>0)
                    ORDER BY cliente.Codigo FOR JSON PATH)
                END,N'[]')) AS OpcionesCliente,
                JSON_QUERY(COALESCE((
                    SELECT TOP (30) proveedor.Ruc AS DocumentoTercero,proveedor.RazonSocial AS NombreTercero
                    FROM dbo.ProveedorHYT AS proveedor
                    WHERE @EsVenta=0 AND proveedor.Activo=1 AND @Busqueda IS NOT NULL
                        AND (@Busqueda=N'' OR CHARINDEX(@Busqueda,proveedor.Ruc)>0 OR CHARINDEX(@Busqueda,proveedor.RazonSocial)>0)
                    ORDER BY proveedor.RazonSocial,proveedor.ProveedorId FOR JSON PATH),N'[]')) AS OpcionesProveedor,
                JSON_QUERY(CASE WHEN @EsVenta=0 OR @ClienteId IS NULL THEN NULL ELSE (
                    SELECT cliente.ClienteId,cliente.NumeroIdentificacion AS DocumentoTercero,
                        CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT_WS(N' ',cliente.Nombres,cliente.Apellidos) ELSE cliente.RazonSocial END AS NombreTercero
                    FROM dbo.Cliente AS cliente WHERE cliente.ClienteId=@ClienteId AND cliente.Activo=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
                END) AS Cliente
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_Consultar TO hyt_app_executor;
GO
