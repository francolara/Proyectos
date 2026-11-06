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
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Reporte_Listar
    @UsuarioId NVARCHAR(128),@Reporte NVARCHAR(40),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@Situacion NVARCHAR(30)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar indicadores o reportes.',1;
        IF @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),@VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END)
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
        LEFT JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1
        WHERE @VerOC=1 AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        IF @Reporte IS NULL OR @Reporte NOT IN(N'GANADOS',N'OC_COMPLETADAS',N'OC_CON_ACTA',N'OC_PENDIENTES_FACTURA') OR @Pagina<1 OR @Pagina>100000000 THROW 51601,N'[HYT_VALIDACION] Reporte o página no válido.',1;
        IF (@Reporte=N'GANADOS' AND @VerPresupuesto=0) OR (@Reporte<>N'GANADOS' AND @VerOC=0) THROW 51600,N'[HYT_PERMISO] No tiene permiso para los documentos del reporte.',1;
        SELECT p.PresupuestoId AS Id,p.NumeroPresupuesto AS Numero,CONVERT(NVARCHAR(100),NULL) AS NumeroCliente,p.FechaRegistro,revision.ClienteNombre,revision.TiendaNombre,revision.CodigoMoneda,revision.Total,revision.Estado,CONVERT(NVARCHAR(100),NULL) AS Guia,CONVERT(NVARCHAR(100),NULL) AS Situacion
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
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
