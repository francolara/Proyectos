-- Firma: FRANCO LARA | Create date: 10/11/2026 | Incluye comprobantes de venta vinculados en los indicadores de facturación.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Indicadores y reportes por cohorte de registro, moneda y permisos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Filtros comerciales, seguimiento y resultados vacíos coherentes de reportes.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual. Guía sin flujo de estados. OC con trabajo completado acredita término para indicadores de facturación. Reporte integral Presupuesto/OC con estados, referencias, registro de relaciones, filtros de pendientes y exportación completa.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Controla la impresión integral por perfil.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Reporte_Listar
    @UsuarioId NVARCHAR(128),@Reporte NVARCHAR(40),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,
    @Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@Situacion NVARCHAR(30)=NULL,
    @Estado NVARCHAR(40)=NULL,@TipoAtencion NVARCHAR(20)=NULL,@Busqueda NVARCHAR(200)=NULL,
    @PendienteGuia BIT=0,@PendienteOC BIT=0,@PendienteConfirmacion BIT=0,@PendienteFacturar BIT=0,@Exportar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Exportar=1 AND dbo.fn_HYTPermiso(@UsuarioId,N'REPORTES.IMPRIMIR')=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para imprimir reportes.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0,@VerGuia BIT=0,@VerConfirmacion BIT=0,@VerFactura BIT=0,@VerEjecucion BIT=0;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar reportes.',1;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END),
            @VerGuia=MAX(CASE WHEN permiso.CodigoPermiso=N'GUIAS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerConfirmacion=MAX(CASE WHEN permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR' THEN 1 ELSE 0 END),
            @VerFactura=MAX(CASE WHEN permiso.CodigoPermiso=N'FACTURADO.CONSULTAR' THEN 1 ELSE 0 END),
            @VerEjecucion=MAX(CASE WHEN permiso.CodigoPermiso=N'EJECUCION.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        DECLARE @Legado NVARCHAR(40)=@Reporte;
        SET @Reporte=CASE WHEN @Reporte=N'GANADOS' THEN N'PRESUPUESTO' WHEN @Reporte IN(N'OC_COMPLETADAS',N'OC_CON_ACTA',N'OC_PENDIENTES_FACTURA') THEN N'OC_CLIENTE' ELSE @Reporte END;
        IF @Reporte IS NULL OR @Reporte NOT IN(N'PRESUPUESTO',N'OC_CLIENTE') OR @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @Desde>=@HastaExclusiva
            THROW 51601,N'[HYT_VALIDACION] Revisa reporte, fechas y página.',1;
        IF (@Reporte=N'PRESUPUESTO' AND @VerPresupuesto=0) OR (@Reporte=N'OC_CLIENTE' AND @VerOC=0)
            OR (@PendienteGuia=1 AND @VerGuia=0) OR (@PendienteOC=1 AND @VerOC=0) OR (@PendienteConfirmacion=1 AND @VerConfirmacion=0) OR (@PendienteFacturar=1 AND @VerFactura=0)
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para los documentos del reporte.',1;
        IF @TipoAtencion IS NOT NULL AND @TipoAtencion NOT IN(N'INVITACION',N'EMERGENCIA') THROW 51601,N'[HYT_VALIDACION] Tipo de atención no válido.',1;
        IF @Estado IS NOT NULL AND ((@Reporte=N'PRESUPUESTO' AND @Estado NOT IN(N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'CANCELADO')) OR (@Reporte=N'OC_CLIENTE' AND @Estado NOT IN(N'REGISTRADA',N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE',N'CANCELADA'))) THROW 51601,N'[HYT_VALIDACION] Estado no válido para el reporte.',1;
        IF @Reporte=N'PRESUPUESTO' AND @Situacion IS NOT NULL THROW 51601,N'[HYT_VALIDACION] El seguimiento corresponde a OC.',1;
        IF @Situacion IS NOT NULL AND @Situacion NOT IN(N'SIN_CONFIRMAR',N'CONFIRMADAS',N'LISTAS',N'FACTURADAS') THROW 51601,N'[HYT_VALIDACION] Situación no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.PresupuestoId ELSE oc.OrdenCompraId END AS Id,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.NumeroPresupuesto ELSE CONVERT(NVARCHAR(100),oc.NumeroRegistroInterno) END AS Numero,
            p.NumeroPresupuesto AS PresupuestoNumero,p.FechaRegistro AS PresupuestoRegistro,p.TipoAtencion,
            actual.ClienteNombre,actual.TiendaNombre,actual.Estado AS EstadoPresupuesto,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Estado ELSE oc.Estado END AS Estado,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.FechaRegistro ELSE oc.FechaRegistro END AS FechaRegistro,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.CodigoMoneda ELSE fuente.CodigoMoneda END AS CodigoMoneda,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Total ELSE COALESCE(oc.Importe,fuente.Total) END AS Total,
            CASE WHEN @VerOC=1 THEN oc.NumeroRegistroInterno END AS OCInterno,CASE WHEN @VerOC=1 THEN oc.NumeroOCCliente END AS NumeroCliente,CASE WHEN @VerOC=1 THEN oc.FechaRegistro END AS OCRegistro,
            CASE WHEN @VerOC=1 THEN oc.Estado END AS EstadoOC,
            CASE WHEN @VerGuia=1 THEN guia.NumeroActa END AS Guia,CASE WHEN @VerGuia=1 AND guia.GuiaId IS NOT NULL THEN N'Guía' END AS TipoGuia,
            CASE WHEN @VerGuia=1 THEN guia.FechaRegistro END AS GuiaRegistro,CASE WHEN @VerGuia=1 THEN guia.FechaActa END AS GuiaFecha,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.EjecucionId END AS EjecucionId,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaRegistro END AS EjecucionRegistro,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaInicioReal END AS EjecucionInicio,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaTerminoReal END AS EjecucionTermino,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.NumeroConfirmacion END AS Confirmacion,CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaRegistro END AS ConfirmacionRegistro,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaConfirmacion END AS ConfirmacionFecha,
            CASE WHEN @VerFactura=1 THEN referencia.CodigoTipoDocumento END AS TipoComprobante,CASE WHEN @VerFactura=1 THEN referencia.Serie END AS Serie,
            CASE WHEN @VerFactura=1 THEN referencia.NumeroDocumento END AS NumeroComprobante,CASE WHEN @VerFactura=1 THEN referencia.FechaRegistro END AS FacturadoRegistro,
            CASE WHEN @VerFactura=1 THEN referencia.FechaDocumento END AS ComprobanteFecha,
            CONVERT(BIT,CASE WHEN @VerFactura=1 AND referencia.ReferenciaId IS NOT NULL THEN 1 ELSE 0 END) AS Facturado,
            CASE WHEN referencia.ReferenciaId IS NOT NULL THEN N'FACTURADA' WHEN confirmacion.ConfirmacionId IS NULL THEN N'SIN CONFIRMAR'
                 WHEN oc.Estado NOT IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') AND (ejecucion.Estado IS NULL OR ejecucion.Estado<>N'TERMINADA') THEN N'TRABAJO PENDIENTE'
                 WHEN guia.GuiaId IS NULL THEN N'GUÍA PENDIENTE' ELSE N'LISTA PARA MARCAR' END AS Situacion
        INTO #Reporte
        FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.RevisionId,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS actual
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        OUTER APPLY (
            SELECT TOP (1) origen.ReferenciaId,origen.CodigoTipoDocumento,origen.Serie,origen.NumeroDocumento,origen.FechaDocumento,origen.FechaRegistro
            FROM (
                SELECT r.ReferenciaId,r.CodigoTipoDocumento,r.Serie,r.NumeroDocumento,r.FechaDocumento,r.FechaRegistro
                FROM dbo.ReferenciaComprobanteExterno AS r WHERE r.OrdenCompraId=oc.OrdenCompraId AND r.Vigente=1
                UNION ALL
                SELECT -c.ComprobanteId,c.Tipo,c.Serie,c.Numero,c.Fecha,c.FechaRegistro
                FROM dbo.ComprobanteHYTPresupuesto AS v INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=v.ComprobanteId
                WHERE v.PresupuestoId=p.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN('01','03') AND c.Serie IS NOT NULL
            ) AS origen ORDER BY origen.FechaRegistro DESC,origen.ReferenciaId
        ) AS referencia
        WHERE ((@Reporte=N'PRESUPUESTO' AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva) OR (@Reporte=N'OC_CLIENTE' AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva))
            AND (@Estado IS NULL OR (@Reporte=N'PRESUPUESTO' AND actual.Estado=@Estado) OR (@Reporte=N'OC_CLIENTE' AND oc.Estado=@Estado))
            AND (@TipoAtencion IS NULL OR p.TipoAtencion=@TipoAtencion)
            AND (@Cliente IS NULL OR actual.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR actual.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR actual.ClienteNombre LIKE @Patron ESCAPE N'~' OR actual.TiendaNombre LIKE @Patron ESCAPE N'~' OR (@VerOC=1 AND oc.NumeroOCCliente LIKE @Patron ESCAPE N'~') OR (@VerGuia=1 AND guia.NumeroActa LIKE @Patron ESCAPE N'~') OR (@VerConfirmacion=1 AND confirmacion.NumeroConfirmacion LIKE @Patron ESCAPE N'~'))
            AND (@PendienteGuia=0 OR guia.GuiaId IS NULL) AND (@PendienteOC=0 OR oc.OrdenCompraId IS NULL)
            AND (@PendienteConfirmacion=0 OR (confirmacion.ConfirmacionId IS NULL AND oc.Estado<>N'CANCELADA')) AND (@PendienteFacturar=0 OR (referencia.ReferenciaId IS NULL AND oc.Estado<>N'CANCELADA'))
            AND (@Legado<>N'GANADOS' OR actual.Estado=N'APROBADO')
            AND (@Legado<>N'OC_COMPLETADAS' OR (oc.Estado<>N'CANCELADA' AND (oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') OR ejecucion.Estado=N'TERMINADA')))
            AND (@Legado<>N'OC_CON_ACTA' OR (oc.Estado<>N'CANCELADA' AND guia.GuiaId IS NOT NULL))
            AND (@Legado<>N'OC_PENDIENTES_FACTURA' OR (oc.Estado<>N'CANCELADA' AND referencia.ReferenciaId IS NULL));
        DELETE reporte FROM #Reporte AS reporte WHERE @Situacion IS NOT NULL AND NOT
            ((@Situacion=N'SIN_CONFIRMAR' AND reporte.Situacion=N'SIN CONFIRMAR') OR (@Situacion=N'CONFIRMADAS' AND reporte.Situacion IN(N'TRABAJO PENDIENTE',N'GUÍA PENDIENTE',N'LISTA PARA MARCAR')) OR (@Situacion=N'LISTAS' AND reporte.Situacion=N'LISTA PARA MARCAR') OR (@Situacion=N'FACTURADAS' AND reporte.Situacion=N'FACTURADA'));
        SELECT COALESCE((SELECT reporte.Id,reporte.Numero,reporte.PresupuestoNumero,reporte.PresupuestoRegistro,reporte.TipoAtencion,reporte.ClienteNombre,reporte.TiendaNombre,reporte.EstadoPresupuesto,reporte.Estado,reporte.FechaRegistro,reporte.CodigoMoneda,reporte.Total,
            reporte.OCInterno,reporte.NumeroCliente,reporte.OCRegistro,reporte.EstadoOC,reporte.Guia,reporte.TipoGuia,reporte.GuiaRegistro,reporte.GuiaFecha,
            reporte.EjecucionId,reporte.EjecucionRegistro,reporte.EjecucionInicio,reporte.EjecucionTermino,reporte.Confirmacion,reporte.ConfirmacionRegistro,reporte.ConfirmacionFecha,
            reporte.TipoComprobante,reporte.Serie,reporte.NumeroComprobante,reporte.FacturadoRegistro,reporte.ComprobanteFecha,reporte.Facturado,
            CASE WHEN @VerFactura=1 AND @VerConfirmacion=1 THEN reporte.Situacion END AS Situacion
        FROM #Reporte AS reporte ORDER BY reporte.FechaRegistro DESC,reporte.Id DESC OFFSET (CASE WHEN @Exportar=1 THEN 0 ELSE (CONVERT(BIGINT,@Pagina)-1)*20 END) ROWS FETCH NEXT (CASE WHEN @Exportar=1 THEN 2147483647 ELSE 20 END) ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Reporte;
        SELECT reporte.CodigoMoneda,SUM(reporte.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Reporte AS reporte GROUP BY reporte.CodigoMoneda ORDER BY reporte.CodigoMoneda;
        DROP TABLE #Reporte;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
