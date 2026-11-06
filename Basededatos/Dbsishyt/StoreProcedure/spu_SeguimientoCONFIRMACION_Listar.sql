-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista Confirmación del cliente por fecha/búsqueda, adjuntos y confirmación; veinte por página.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.ConfirmacionId AS Id,documento.OrdenCompraId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) AS NumeroInterno,CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) AS NumeroDocumento,CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END AS Estado,documento.FechaConfirmacion AS FechaDocumento,
            fuente.CodigoMoneda,CONVERT(DECIMAL(19,2),NULL) AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),oc.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=documento.ConfirmacionId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=documento.ConfirmacionId) AS PrimerArchivoId,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,CONVERT(NVARCHAR(MAX),NULL) AS Facturado
        INTO #Pagina FROM dbo.ConfirmacionFacturacion AS documento INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=documento.OrdenCompraId
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.ConfirmacionId=documento.ConfirmacionId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END=@Estado)
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
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
