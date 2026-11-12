-- Firma: FRANCO LARA | Create date: 10/11/2026 | Incluye la descripción de las partidas de la última revisión en la búsqueda y la facturación por comprobante vinculado.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista presupuestos por búsqueda y rango UTC; última revisión, OC y guía, veinte por página.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual; incluye EjecucionId para identificar documentos registrados en las acciones. Filtros tipo/pendientes y KPI por estado del conjunto completo.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Agrega referencias de comprobantes del presupuesto con permiso de facturación.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),
    @Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(20)=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@TipoAtencion NVARCHAR(20)=NULL,@PendienteGuia BIT=0,@PendienteOC BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde
            THROW 51301,N'[HYT_VALIDACION] Revisa página y rango de fechas.',1;
        IF @TipoAtencion IS NOT NULL AND @TipoAtencion NOT IN(N'INVITACION',N'EMERGENCIA') THROW 51301,N'[HYT_VALIDACION] Tipo de atención no válido.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.Descripcion,p.TipoAtencion,p.ClienteId,p.TiendaId,
            revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,
            revision.CodigoMoneda,revision.Subtotal,revision.Impuesto,revision.Total,revision.FechaAprobacion,
            p.FechaRegistro,creador.Nombres AS UsuarioRegistro,p.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,
            oc.OrdenCompraId,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,
            CONVERT(BIT,CASE WHEN dbo.fn_HYTPermiso(@UsuarioId,N'FACTURADO.CONSULTAR')=1 AND (
                EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPresupuesto AS v INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=v.ComprobanteId WHERE v.PresupuestoId=p.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN('01','03') AND c.Serie IS NOT NULL)
                OR EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS factura WHERE factura.OrdenCompraId=oc.OrdenCompraId AND factura.Vigente=1)) THEN 1 ELSE 0 END) AS Facturado,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia
        INTO #Pagina FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Subtotal,r.Impuesto,r.Total,r.FechaAprobacion,r.FechaActualizacion,r.UsuarioActualizacion
            FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        WHERE p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva AND (@TipoAtencion IS NULL OR p.TipoAtencion=@TipoAtencion) AND (@PendienteGuia=0 OR guia.GuiaId IS NULL) AND (@PendienteOC=0 OR oc.OrdenCompraId IS NULL)
          AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR p.Descripcion LIKE @Patron ESCAPE N'~' OR revision.ClienteNombre LIKE @Patron ESCAPE N'~' OR revision.TiendaNombre LIKE @Patron ESCAPE N'~'
                OR EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS partida
                    WHERE partida.RevisionId=revision.RevisionId AND partida.Descripcion LIKE @Patron ESCAPE N'~'));
        SELECT pagina.PresupuestoId,pagina.NumeroInterno,pagina.Descripcion,pagina.TipoAtencion,pagina.ClienteId,pagina.TiendaId,
            pagina.RevisionId,pagina.NumeroRevision,pagina.ClienteNombre,pagina.TiendaNombre,pagina.Estado,pagina.CodigoMoneda,pagina.Subtotal,pagina.Impuesto,pagina.Total,pagina.FechaAprobacion,
            pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.OrdenCompraId,pagina.NumeroOCCliente,pagina.GuiaId,pagina.NumeroActa,pagina.EstadoGuia,pagina.AdjuntosGuia,pagina.EjecucionId,pagina.Facturado,
            CASE WHEN pagina.Facturado=1 THEN (
                SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), comprobantes.Numero),N', ') WITHIN GROUP (ORDER BY comprobantes.Numero)
                FROM (
                    SELECT CONCAT(c.Serie,N'-',c.Numero) AS Numero
                    FROM dbo.ComprobanteHYTPresupuesto AS vinculo
                    INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=vinculo.ComprobanteId
                    WHERE vinculo.PresupuestoId=pagina.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN ('01','03') AND c.Serie IS NOT NULL
                    UNION
                    SELECT CONCAT(referencia.Serie,N'-',referencia.NumeroDocumento)
                    FROM dbo.ReferenciaComprobanteExterno AS referencia
                    WHERE referencia.OrdenCompraId=pagina.OrdenCompraId AND referencia.Vigente=1
                ) AS comprobantes
            ) END AS Comprobantes
        FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado) ORDER BY pagina.FechaRegistro DESC,pagina.PresupuestoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado);
        SELECT pagina.Estado,COUNT_BIG(1) AS Cantidad FROM #Pagina AS pagina GROUP BY pagina.Estado;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
