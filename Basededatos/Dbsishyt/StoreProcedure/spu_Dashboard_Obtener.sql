-- Firma: FRANCO LARA | Create date: 10/11/2026 | Incluye comprobantes de venta vinculados en los indicadores de facturación. Consolida los eventos de facturación por OC sin duplicados.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Indicadores y reportes por cohorte de registro, moneda y permisos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Filtros comerciales, primeros eventos, resultados vacíos y pendientes históricos paginados.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual. Guía sin flujo de estados. OC con trabajo completado acredita término para indicadores de facturación. Panel operativo por estados actuales y pendientes de OC, guía, confirmación y facturación. Curvas de Confirmación del cliente y Facturado por registro vigente; elimina cálculo de primeros eventos históricos.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Resume la cobertura de guías y OC vigentes en presupuestos aceptados según permisos.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Dashboard_Obtener
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@PaginaPendientes INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar indicadores o reportes.',1;
        IF @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0,@VerGuia BIT=0;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),@VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END),
            @VerGuia=MAX(CASE WHEN permiso.CodigoPermiso=N'GUIAS.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.FechaRegistro,p.TipoAtencion,revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,revision.CodigoMoneda,revision.Total
        INTO #Presupuestos FROM dbo.Presupuesto AS p CROSS APPLY(SELECT TOP(1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        WHERE @VerPresupuesto=1 AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        SELECT oc.OrdenCompraId,oc.PresupuestoId,oc.NumeroRegistroInterno,oc.NumeroOCCliente,oc.FechaRegistro,oc.Estado,revision.CodigoMoneda,revision.Total,revision.ClienteNombre,revision.TiendaNombre,p.NumeroPresupuesto AS PresupuestoNumero,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.Estado AS EstadoEjecucion,confirmacion.ConfirmacionId,referencia.ReferenciaId
        INTO #Ordenes FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=oc.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=oc.PresupuestoId
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
        WHERE @VerOC=1 AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        IF @PaginaPendientes<1 OR @PaginaPendientes>100000000 THROW 51601,N'[HYT_VALIDACION] Página no válida.',1;
        DECLARE @VerPendientes BIT=0;
        IF @VerPresupuesto=1 AND @VerOC=1 AND @VerGuia=1 SET @VerPendientes=1;
        SELECT p.PresupuestoId AS Id,p.NumeroPresupuesto AS Numero,p.FechaRegistro,p.TipoAtencion,revision.ClienteNombre,revision.TiendaNombre,
            CONVERT(BIT,CASE WHEN oc.OrdenCompraId IS NULL THEN 1 ELSE 0 END) AS SinOC,CONVERT(BIT,CASE WHEN guia.GuiaId IS NULL THEN 1 ELSE 0 END) AS SinGuia,
            CONVERT(BIT,CASE WHEN p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO' THEN 1 ELSE 0 END) AS EmergenciaPorRegularizar,
            DATEDIFF(DAY,CONVERT(DATE,DATEADD(HOUR,-5,p.FechaRegistro)),CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()))) AS DiasPendiente
        INTO #Pendientes FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.ClienteNombre,r.TiendaNombre,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId AND oc.Estado<>N'CANCELADA'
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        WHERE @VerPendientes=1 AND revision.Estado<>N'CANCELADO' AND p.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (oc.OrdenCompraId IS NULL OR guia.GuiaId IS NULL OR (p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO'));
        SELECT evento.FechaRegistro,evento.Tipo INTO #Eventos
        FROM (SELECT confirmacion.OrdenCompraId,confirmacion.FechaRegistro,N'CONFIRMACION' AS Tipo FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.Vigente=1
            UNION ALL SELECT ocFactura.OrdenCompraId,MAX(origen.FechaRegistro),N'FACTURADO'
            FROM dbo.OrdenCompraCliente AS ocFactura
            CROSS APPLY (
                SELECT r.FechaRegistro FROM dbo.ReferenciaComprobanteExterno AS r WHERE r.OrdenCompraId=ocFactura.OrdenCompraId AND r.Vigente=1
                UNION ALL
                SELECT c.FechaRegistro FROM dbo.ComprobanteHYTPresupuesto AS v INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=v.ComprobanteId
                WHERE v.PresupuestoId=ocFactura.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN('01','03') AND c.Serie IS NOT NULL
            ) AS origen GROUP BY ocFactura.OrdenCompraId) AS evento
        INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=evento.OrdenCompraId
        INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        WHERE @VerOC=1 AND oc.Estado<>N'CANCELADA' AND evento.FechaRegistro>=@Desde AND evento.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        SELECT (SELECT @VerPresupuesto AS VerPresupuestos,@VerOC AS VerOC,@VerPendientes AS VerPendientes,SYSUTCDATETIME() AS ActualizadoUTC,@Desde AS DesdeUTC,@HastaExclusiva AS HastaExclusivaUTC,
            (SELECT COUNT_BIG(1) FROM #Pendientes) AS TotalPendientes,
            (SELECT COUNT_BIG(1) FROM #Pendientes AS pendiente WHERE pendiente.FechaRegistro>=@Desde) AS PendientesPeriodo,
            JSON_QUERY(COALESCE((SELECT pendiente.Id,pendiente.Numero,pendiente.FechaRegistro,pendiente.TipoAtencion,pendiente.ClienteNombre,pendiente.TiendaNombre,pendiente.SinOC,pendiente.SinGuia,pendiente.EmergenciaPorRegularizar,pendiente.DiasPendiente FROM #Pendientes AS pendiente ORDER BY pendiente.FechaRegistro,pendiente.Id OFFSET (CONVERT(BIGINT,@PaginaPendientes)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH),N'[]')) AS Pendientes,
            JSON_QUERY(COALESCE((SELECT presupuesto.Estado,COUNT_BIG(1) AS Cantidad FROM #Presupuestos AS presupuesto GROUP BY presupuesto.Estado FOR JSON PATH),N'[]')) AS EstadosPresupuesto,
            JSON_QUERY(COALESCE((SELECT orden.Estado,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden GROUP BY orden.Estado FOR JSON PATH),N'[]')) AS EstadosOC,
            (SELECT COUNT_BIG(1) FROM #Presupuestos) AS Presupuestos,
            (SELECT COUNT_BIG(1) FROM #Ordenes) AS Ordenes,
            JSON_QUERY((SELECT
                CASE WHEN @VerGuia=1 THEN COUNT_BIG(guia.GuiaId) END AS ConGuia,
                CASE WHEN @VerOC=1 THEN COUNT_BIG(oc.OrdenCompraId) END AS ConOC,
                CASE WHEN @VerGuia=1 AND @VerOC=1 THEN COUNT_BIG(CASE WHEN guia.GuiaId IS NOT NULL AND oc.OrdenCompraId IS NOT NULL THEN 1 END) END AS ConAmbos
                FROM #Presupuestos AS presupuesto
                LEFT JOIN dbo.GuiaConformidad AS guia ON @VerGuia=1 AND guia.PresupuestoId=presupuesto.PresupuestoId
                LEFT JOIN dbo.OrdenCompraCliente AS oc ON @VerOC=1 AND oc.PresupuestoId=presupuesto.PresupuestoId AND oc.Estado<>N'CANCELADA'
                WHERE presupuesto.Estado=N'APROBADO'
                FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)) AS CoberturaPresupuesto,
            (SELECT COUNT_BIG(1) FROM #Presupuestos AS presupuesto WHERE NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=presupuesto.PresupuestoId)) AS PresupuestosSinOC,
            (SELECT COUNT_BIG(1) FROM #Presupuestos AS presupuesto WHERE NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=presupuesto.PresupuestoId)) AS PresupuestosSinGuia,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NULL) AS PendientesFacturar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NULL) AS SinConfirmar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NOT NULL AND orden.ReferenciaId IS NULL) AS ConfirmadasSinMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NULL AND orden.ConfirmacionId IS NOT NULL AND (orden.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') OR orden.EstadoEjecucion=N'TERMINADA') AND orden.GuiaId IS NOT NULL) AS ListasParaMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NOT NULL) AS Facturadas,
            JSON_QUERY(COALESCE((SELECT presupuesto.CodigoMoneda,SUM(presupuesto.Total) AS Importe,COUNT_BIG(1) AS Cantidad,SUM(CONVERT(BIGINT,CASE WHEN presupuesto.Total IS NULL THEN 1 ELSE 0 END)) AS PreciosPendientes FROM #Presupuestos AS presupuesto WHERE presupuesto.Estado<>N'CANCELADO' GROUP BY presupuesto.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS ImportesPresupuesto,
            JSON_QUERY(COALESCE((SELECT orden.CodigoMoneda,SUM(orden.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' GROUP BY orden.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS ImportesOC,
            JSON_QUERY(COALESCE((SELECT movimiento.Fecha,SUM(movimiento.Presupuestos) AS Presupuestos,SUM(movimiento.Ordenes) AS Ordenes,SUM(movimiento.Confirmaciones) AS Confirmaciones,SUM(movimiento.Marcados) AS Marcados FROM
                (SELECT CONVERT(DATE,DATEADD(HOUR,-5,presupuesto.FechaRegistro)) AS Fecha,CONVERT(BIGINT,1) AS Presupuestos,CONVERT(BIGINT,0) AS Ordenes,CONVERT(BIGINT,0) AS Confirmaciones,CONVERT(BIGINT,0) AS Marcados FROM #Presupuestos AS presupuesto
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,orden.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,1),CONVERT(BIGINT,0),CONVERT(BIGINT,0) FROM #Ordenes AS orden
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,evento.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,0),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'CONFIRMACION' THEN 1 ELSE 0 END),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'FACTURADO' THEN 1 ELSE 0 END) FROM #Eventos AS evento) AS movimiento
                GROUP BY movimiento.Fecha ORDER BY movimiento.Fecha FOR JSON PATH),N'[]')) AS Movimientos
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) AS Datos;
        DROP TABLE #Presupuestos;DROP TABLE #Ordenes;DROP TABLE #Pendientes;DROP TABLE #Eventos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
