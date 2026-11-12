-- Firma: FRANCO LARA | Create date: 10/11/2026 | Integra comprobantes emitidos y sus OC facturadas.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista Orden de Compra - Cliente por fecha/búsqueda, adjuntos y confirmación; veinte por página.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual; incluye número de Guía y cantidad de adjuntos según permiso GUIAS.CONSULTAR. KPI globales por estado y filtros de confirmación/facturación pendientes.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Agrega fecha de recepción y comprobante emitido para el reporte del listado.
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
            CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) AS NumeroInterno,documento.NumeroOCCliente AS NumeroDocumento,documento.Estado AS Estado,documento.FechaEmisionOC AS FechaDocumento,documento.FechaRecepcionOC,
            fuente.CodigoMoneda,documento.Importe AS Importe,
            emitido.ComprobanteId,emitido.NumeroComprobante,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),p.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS PrimerArchivoId,
            guia.GuiaId,guia.NumeroActa AS NumeroGuia,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.FechaRegistro,autor.Nombres AS UsuarioRegistro FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.AspNetUsers AS autor ON autor.Id=referencia.UsuarioRegistro WHERE referencia.OrdenCompraId=documento.OrdenCompraId AND referencia.Vigente=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permisoFactura ON permisoFactura.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permisoFactura.CodigoPermiso=N'FACTURADO.CONSULTAR') FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Facturado
        INTO #Pagina FROM dbo.OrdenCompraCliente AS documento
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=documento.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=documento.RevisionId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=documento.OrdenCompraId AND confirmacion.Vigente=1
        OUTER APPLY
        (
            SELECT TOP (1) c.ComprobanteId, CONCAT(c.Serie,N'-',c.Numero) AS NumeroComprobante
            FROM dbo.ComprobanteHYTPresupuesto AS vinculo
            INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=vinculo.ComprobanteId
            WHERE vinculo.PresupuestoId=p.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN ('01','03')
                AND c.Serie IS NOT NULL AND dbo.fn_HYTPermiso(@UsuarioId,N'FACTURADO.CONSULTAR')=1
            ORDER BY c.FechaRegistro DESC,c.ComprobanteId DESC
        ) AS emitido
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId AND EXISTS
            (SELECT 1 FROM dbo.AspNetUserRoles AS rolGuia INNER JOIN dbo.RolPermiso AS permisoGuia ON permisoGuia.RoleId=rolGuia.RoleId WHERE rolGuia.UserId=@UsuarioId AND permisoGuia.CodigoPermiso=N'GUIAS.CONSULTAR')
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) LIKE @Patron ESCAPE N'~' OR documento.NumeroOCCliente LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL))
            AND (@PendienteConfirmacion=0 OR confirmacion.ConfirmacionId IS NULL)
            AND (@PendienteFacturar=0 OR NOT EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS pendiente WHERE pendiente.OrdenCompraId=documento.OrdenCompraId AND pendiente.Vigente=1)
                AND NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPresupuesto AS v INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=v.ComprobanteId
                    WHERE v.PresupuestoId=p.PresupuestoId AND c.EsVenta=1 AND c.Tipo IN('01','03') AND c.Serie IS NOT NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.GuiaId,pagina.NumeroGuia,pagina.AdjuntosGuia,pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,pagina.ComprobanteId,pagina.NumeroComprobante,pagina.FechaRecepcionOC,JSON_QUERY(pagina.Facturado) AS Facturado
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
