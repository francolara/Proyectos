"""Genera dashboard y reportes HTML por consultas SQL; no conecta con el motor."""
from pathlib import Path
import json
BASE=Path(__file__).resolve().parents[3]/'Basededatos/Dbsishyt'
def save(path,text):path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
def proc(name,params,body):
    text=f'''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Indicadores y reportes por cohorte de registro, moneda y permisos.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.{name}
{params}
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
'''+body+'''    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
'''
    source='StoreProcedure/'+name+'.sql'
    if name in ('spu_Dashboard_Obtener','spu_Reporte_Listar'):
        description='Filtros comerciales, primeros eventos, resultados vacíos y pendientes históricos paginados.' if name=='spu_Dashboard_Obtener' else 'Filtros comerciales, seguimiento y resultados vacíos coherentes de reportes.'
        text=text.replace('USE [SistemaWebHYT];',f'-- =============================================\n-- Author:        FRANCO LARA\n-- Create date:   05/11/2026\n-- Description:   {description}\n-- =============================================\nUSE [SistemaWebHYT];',1)
    save(BASE/source,text);return source
permission='''        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar indicadores o reportes.',1;
        IF @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),@VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
'''
cohorts='''        SELECT p.PresupuestoId,p.NumeroInterno,p.FechaRegistro,p.TipoAtencion,revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,revision.CodigoMoneda,revision.Total
        INTO #Presupuestos FROM dbo.Presupuesto AS p CROSS APPLY(SELECT TOP(1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        WHERE @VerPresupuesto=1 AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva;
        SELECT oc.OrdenCompraId,oc.PresupuestoId,oc.NumeroRegistroInterno,oc.NumeroOCCliente,oc.FechaRegistro,oc.Estado,revision.CodigoMoneda,revision.Total,revision.ClienteNombre,revision.TiendaNombre,p.NumeroInterno AS PresupuestoNumero,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.Estado AS EstadoEjecucion,confirmacion.ConfirmacionId,referencia.ReferenciaId
        INTO #Ordenes FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=oc.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=oc.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1
        WHERE @VerOC=1 AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva;
'''
dashboard_cohorts=cohorts.replace('p.FechaRegistro<@HastaExclusiva;',"p.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');").replace('oc.FechaRegistro<@HastaExclusiva;',"oc.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');")
dashboard_extra='''        IF @PaginaPendientes<1 OR @PaginaPendientes>100000000 THROW 51601,N'[HYT_VALIDACION] Página no válida.',1;
        DECLARE @VerPendientes BIT=0;
        IF @VerPresupuesto=1 AND @VerOC=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'GUIAS.CONSULTAR') SET @VerPendientes=1;
        SELECT p.PresupuestoId AS Id,p.NumeroInterno AS Numero,p.FechaRegistro,p.TipoAtencion,revision.ClienteNombre,revision.TiendaNombre,
            CONVERT(BIT,CASE WHEN oc.OrdenCompraId IS NULL THEN 1 ELSE 0 END) AS SinOC,CONVERT(BIT,CASE WHEN guia.GuiaId IS NULL OR guia.Estado=N'ANULADA' THEN 1 ELSE 0 END) AS SinGuia,
            CONVERT(BIT,CASE WHEN p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO' THEN 1 ELSE 0 END) AS EmergenciaPorRegularizar,
            DATEDIFF(DAY,CONVERT(DATE,DATEADD(HOUR,-5,p.FechaRegistro)),CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()))) AS DiasPendiente
        INTO #Pendientes FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.ClienteNombre,r.TiendaNombre,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId AND oc.Estado<>N'CANCELADA'
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        WHERE @VerPendientes=1 AND revision.Estado<>N'CANCELADO' AND p.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (oc.OrdenCompraId IS NULL OR guia.GuiaId IS NULL OR guia.Estado=N'ANULADA' OR (p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO'));
        SELECT evento.OrdenCompraId,MIN(evento.FechaRegistro) AS FechaRegistro,evento.Tipo
        INTO #Eventos FROM
            (SELECT confirmacion.OrdenCompraId,confirmacion.FechaRegistro,N'CONFIRMACION' AS Tipo FROM dbo.ConfirmacionFacturacion AS confirmacion
             UNION ALL SELECT referencia.OrdenCompraId,referencia.FechaRegistro,N'MARCADO' FROM dbo.ReferenciaComprobanteExterno AS referencia) AS evento
        INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=evento.OrdenCompraId INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        WHERE @VerOC=1 AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
        GROUP BY evento.OrdenCompraId,evento.Tipo;
'''
body=permission+dashboard_cohorts+dashboard_extra+'''        SELECT (SELECT @VerPresupuesto AS VerPresupuestos,@VerOC AS VerOC,@VerPendientes AS VerPendientes,SYSUTCDATETIME() AS ActualizadoUTC,@Desde AS DesdeUTC,@HastaExclusiva AS HastaExclusivaUTC,
            (SELECT COUNT_BIG(1) FROM #Pendientes) AS TotalPendientes,
            (SELECT COUNT_BIG(1) FROM #Pendientes AS pendiente WHERE pendiente.FechaRegistro>=@Desde) AS PendientesPeriodo,
            JSON_QUERY((SELECT pendiente.Id,pendiente.Numero,pendiente.FechaRegistro,pendiente.TipoAtencion,pendiente.ClienteNombre,pendiente.TiendaNombre,pendiente.SinOC,pendiente.SinGuia,pendiente.EmergenciaPorRegularizar,pendiente.DiasPendiente FROM #Pendientes AS pendiente ORDER BY pendiente.FechaRegistro,pendiente.Id OFFSET (CONVERT(BIGINT,@PaginaPendientes)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH)) AS Pendientes,
            JSON_QUERY((SELECT presupuesto.Estado,COUNT_BIG(1) AS Cantidad FROM #Presupuestos AS presupuesto GROUP BY presupuesto.Estado FOR JSON PATH)) AS EstadosPresupuesto,
            JSON_QUERY((SELECT orden.Estado,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden GROUP BY orden.Estado FOR JSON PATH)) AS EstadosOC,
            (SELECT COUNT_BIG(1) FROM #Presupuestos) AS Presupuestos,
            (SELECT COUNT_BIG(1) FROM #Ordenes) AS Ordenes,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NULL) AS SinConfirmar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NOT NULL AND orden.ReferenciaId IS NULL) AS ConfirmadasSinMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NULL AND orden.ConfirmacionId IS NOT NULL AND orden.EstadoEjecucion=N'TERMINADA' AND orden.EstadoGuia=N'ACEPTADA') AS ListasParaMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NOT NULL) AS Facturadas,
            JSON_QUERY((SELECT presupuesto.CodigoMoneda,SUM(presupuesto.Total) AS Importe,COUNT_BIG(1) AS Cantidad,SUM(CONVERT(BIGINT,CASE WHEN presupuesto.Total IS NULL THEN 1 ELSE 0 END)) AS PreciosPendientes FROM #Presupuestos AS presupuesto WHERE presupuesto.Estado<>N'CANCELADO' GROUP BY presupuesto.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES)) AS ImportesPresupuesto,
            JSON_QUERY((SELECT orden.CodigoMoneda,SUM(orden.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' GROUP BY orden.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES)) AS ImportesOC,
            JSON_QUERY((SELECT movimiento.Fecha,SUM(movimiento.Presupuestos) AS Presupuestos,SUM(movimiento.Ordenes) AS Ordenes,SUM(movimiento.Confirmaciones) AS Confirmaciones,SUM(movimiento.Marcados) AS Marcados FROM
                (SELECT CONVERT(DATE,DATEADD(HOUR,-5,presupuesto.FechaRegistro)) AS Fecha,CONVERT(BIGINT,1) AS Presupuestos,CONVERT(BIGINT,0) AS Ordenes,CONVERT(BIGINT,0) AS Confirmaciones,CONVERT(BIGINT,0) AS Marcados FROM #Presupuestos AS presupuesto
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,orden.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,1),CONVERT(BIGINT,0),CONVERT(BIGINT,0) FROM #Ordenes AS orden
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,evento.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,0),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'CONFIRMACION' THEN 1 ELSE 0 END),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'MARCADO' THEN 1 ELSE 0 END) FROM #Eventos AS evento WHERE evento.FechaRegistro>=@Desde AND evento.FechaRegistro<@HastaExclusiva) AS movimiento
                GROUP BY movimiento.Fecha ORDER BY movimiento.Fecha FOR JSON PATH)) AS Movimientos
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) AS Datos;
        DROP TABLE #Presupuestos;DROP TABLE #Ordenes;DROP TABLE #Pendientes;DROP TABLE #Eventos;
'''
body=body.replace('JSON_QUERY((SELECT','JSON_QUERY(COALESCE((SELECT').replace('FOR JSON PATH)) AS',"FOR JSON PATH),N'[]')) AS").replace('FOR JSON PATH,INCLUDE_NULL_VALUES)) AS',"FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS")
sources=[proc('spu_Dashboard_Obtener','    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@PaginaPendientes INT=1',body)]
body=permission+dashboard_cohorts+'''        IF @Reporte IS NULL OR @Reporte NOT IN(N'GANADOS',N'OC_COMPLETADAS',N'OC_CON_ACTA',N'OC_PENDIENTES_FACTURA') OR @Pagina<1 OR @Pagina>100000000 THROW 51601,N'[HYT_VALIDACION] Reporte o página no válido.',1;
        IF (@Reporte=N'GANADOS' AND @VerPresupuesto=0) OR (@Reporte<>N'GANADOS' AND @VerOC=0) THROW 51600,N'[HYT_PERMISO] No tiene permiso para los documentos del reporte.',1;
        SELECT p.PresupuestoId AS Id,p.NumeroInterno AS Numero,CONVERT(NVARCHAR(100),NULL) AS NumeroCliente,p.FechaRegistro,revision.ClienteNombre,revision.TiendaNombre,revision.CodigoMoneda,revision.Total,revision.Estado,CONVERT(NVARCHAR(100),NULL) AS Guia,CONVERT(NVARCHAR(100),NULL) AS Situacion
        INTO #Reporte FROM dbo.Presupuesto AS p CROSS APPLY(SELECT TOP(1) r.ClienteNombre,r.TiendaNombre,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId AND r.Estado=N'APROBADO' ORDER BY r.NumeroRevision DESC) AS revision
        WHERE @Reporte=N'GANADOS' AND @VerPresupuesto=1 AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND NOT EXISTS(SELECT 1 FROM #Presupuestos AS actual WHERE actual.PresupuestoId=p.PresupuestoId AND actual.Estado=N'CANCELADO');
        INSERT INTO #Reporte(Id,Numero,NumeroCliente,FechaRegistro,ClienteNombre,TiendaNombre,CodigoMoneda,Total,Estado,Guia,Situacion)
        SELECT orden.OrdenCompraId,orden.NumeroRegistroInterno,orden.NumeroOCCliente,orden.FechaRegistro,orden.ClienteNombre,orden.TiendaNombre,orden.CodigoMoneda,orden.Total,orden.Estado,orden.NumeroActa,
            CASE WHEN orden.ReferenciaId IS NOT NULL THEN N'FACTURADA' WHEN orden.ConfirmacionId IS NULL THEN N'SIN CONFIRMAR' WHEN orden.EstadoEjecucion<>N'TERMINADA' OR orden.EstadoEjecucion IS NULL THEN N'TRABAJO PENDIENTE' WHEN orden.EstadoGuia<>N'ACEPTADA' OR orden.EstadoGuia IS NULL THEN N'GUÍA PENDIENTE' ELSE N'LISTA PARA MARCAR' END
        FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND ((@Reporte=N'OC_COMPLETADAS' AND orden.EstadoEjecucion=N'TERMINADA') OR (@Reporte=N'OC_CON_ACTA' AND orden.GuiaId IS NOT NULL AND orden.EstadoGuia<>N'ANULADA') OR (@Reporte=N'OC_PENDIENTES_FACTURA' AND orden.ReferenciaId IS NULL));
        IF @Reporte=N'GANADOS' AND @Situacion IS NOT NULL THROW 51601,N'[HYT_VALIDACION] El seguimiento corresponde a reportes OC.',1;
        IF @Situacion IS NOT NULL AND @Situacion NOT IN(N'SIN_CONFIRMAR',N'CONFIRMADAS',N'LISTAS',N'FACTURADAS') THROW 51601,N'[HYT_VALIDACION] Situación no válida.',1;
        DELETE reporte FROM #Reporte AS reporte WHERE @Situacion IS NOT NULL AND NOT
            ((@Situacion=N'SIN_CONFIRMAR' AND reporte.Situacion=N'SIN CONFIRMAR') OR (@Situacion=N'CONFIRMADAS' AND reporte.Situacion IN(N'TRABAJO PENDIENTE',N'GUÍA PENDIENTE',N'LISTA PARA MARCAR')) OR (@Situacion=N'LISTAS' AND reporte.Situacion=N'LISTA PARA MARCAR') OR (@Situacion=N'FACTURADAS' AND reporte.Situacion=N'FACTURADA'));
        SELECT COALESCE((SELECT reporte.Id,reporte.Numero,reporte.NumeroCliente,reporte.FechaRegistro,reporte.ClienteNombre,reporte.TiendaNombre,reporte.CodigoMoneda,reporte.Total,reporte.Estado,reporte.Guia,reporte.Situacion FROM #Reporte AS reporte ORDER BY reporte.FechaRegistro DESC,reporte.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Reporte;
        SELECT reporte.CodigoMoneda,SUM(reporte.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Reporte AS reporte GROUP BY reporte.CodigoMoneda ORDER BY reporte.CodigoMoneda;
        DROP TABLE #Reporte;DROP TABLE #Presupuestos;DROP TABLE #Ordenes;
'''
sources.append(proc('spu_Reporte_Listar','    @UsuarioId NVARCHAR(128),@Reporte NVARCHAR(40),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@Situacion NVARCHAR(30)=NULL',body))
grant='Script/019_PermisosEtapa8.sql'
save(BASE/grant,'-- =============================================\n-- Author:        FRANCO LARA\n-- Create date:   03/11/2026\n-- Description:   Permisos limitados de dashboard y reportes.\n-- =============================================\nUSE [SistemaWebHYT];\nGO\n'+''.join(f'GRANT EXECUTE ON OBJECT::dbo.{Path(p).stem} TO hyt_app_executor;\n' for p in sources)+'GO\n');sources.append(grant)
save(BASE/'manifest-etapa8.json',json.dumps({'database':'SistemaWebHYT','requires':'etapa7','scripts':sources},indent=2)+'\n')
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))
manifest['scripts']=[p for n in range(1,10) if (BASE/f'manifest-etapa{n}.json').exists() for p in json.loads((BASE/f'manifest-etapa{n}.json').read_text(encoding='utf-8'))['scripts']]
save(BASE/'manifest.json',json.dumps(manifest,indent=2)+'\n')
print('Dashboard/reportes: tres fuentes preparadas; sin ejecución SQL desde generador.')
