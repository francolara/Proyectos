-- Firma: FRANCO LARA | Create date: 06/11/2026 | KPI por estado y filtros de pendientes en listados de presupuesto y OC.

-- Fuente: StoreProcedure/spu_Presupuesto_Listar.sql
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
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR p.Descripcion LIKE @Patron ESCAPE N'~' OR revision.ClienteNombre LIKE @Patron ESCAPE N'~' OR revision.TiendaNombre LIKE @Patron ESCAPE N'~');
        SELECT pagina.PresupuestoId,pagina.NumeroInterno,pagina.Descripcion,pagina.TipoAtencion,pagina.ClienteId,pagina.TiendaId,
            pagina.RevisionId,pagina.NumeroRevision,pagina.ClienteNombre,pagina.TiendaNombre,pagina.Estado,pagina.CodigoMoneda,pagina.Subtotal,pagina.Impuesto,pagina.Total,pagina.FechaAprobacion,
            pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.OrdenCompraId,pagina.NumeroOCCliente,pagina.GuiaId,pagina.NumeroActa,pagina.EstadoGuia,pagina.AdjuntosGuia,pagina.EjecucionId
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

-- Fuente: StoreProcedure/spu_SeguimientoOC_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@PendienteConfirmacion BIT=0,@PendienteFacturar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PendienteFacturar=1 AND NOT EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'FACTURADO.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para consultar facturación pendiente.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.OrdenCompraId AS Id,documento.PresupuestoId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) AS NumeroInterno,documento.NumeroOCCliente AS NumeroDocumento,documento.Estado AS Estado,documento.FechaEmisionOC AS FechaDocumento,
            fuente.CodigoMoneda,documento.Importe AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),p.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS PrimerArchivoId,
            guia.GuiaId,guia.NumeroActa AS NumeroGuia,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.FechaRegistro,autor.Nombres AS UsuarioRegistro FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.AspNetUsers AS autor ON autor.Id=referencia.UsuarioRegistro WHERE referencia.OrdenCompraId=documento.OrdenCompraId AND referencia.Vigente=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permisoFactura ON permisoFactura.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permisoFactura.CodigoPermiso=N'FACTURADO.CONSULTAR') FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Facturado
        INTO #Pagina FROM dbo.OrdenCompraCliente AS documento
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=documento.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=documento.RevisionId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=documento.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId AND EXISTS
            (SELECT 1 FROM dbo.AspNetUserRoles AS rolGuia INNER JOIN dbo.RolPermiso AS permisoGuia ON permisoGuia.RoleId=rolGuia.RoleId WHERE rolGuia.UserId=@UsuarioId AND permisoGuia.CodigoPermiso=N'GUIAS.CONSULTAR')
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) LIKE @Patron ESCAPE N'~' OR documento.NumeroOCCliente LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL))
            AND (@PendienteConfirmacion=0 OR confirmacion.ConfirmacionId IS NULL)
            AND (@PendienteFacturar=0 OR NOT EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS pendiente WHERE pendiente.OrdenCompraId=documento.OrdenCompraId AND pendiente.Vigente=1));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.GuiaId,pagina.NumeroGuia,pagina.AdjuntosGuia,pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado) ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
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
