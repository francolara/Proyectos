-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Marcado Facturado externo con permisos y trazabilidad.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Facturado_Obtener
    @UsuarioId NVARCHAR(128),@OrdenCompraId BIGINT,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'FACTURADO.CONSULTAR')
            THROW 51500,N'[HYT_PERMISO] No tiene permiso para esta referencia de facturación.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51502,N'[HYT_VALIDACION] Página no válida.',1;
        SELECT (SELECT oc.OrdenCompraId,oc.NumeroRegistroInterno,oc.NumeroOCCliente,oc.PresupuestoId,oc.Estado,oc.VersionFila,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT referencia.ReferenciaId,referencia.ConfirmacionId,referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.VersionFila,referencia.FechaRegistro,autor.Nombres AS UsuarioRegistro,referencia.FechaActualizacion,editor.Nombres AS UsuarioActualizacion FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.AspNetUsers AS autor ON autor.Id=referencia.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=referencia.UsuarioActualizacion WHERE referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Referencia,
            JSON_QUERY((SELECT referencia.ReferenciaId,referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.Vigente,referencia.MotivoRetiro FROM dbo.ReferenciaComprobanteExterno AS referencia WHERE referencia.OrdenCompraId=oc.OrdenCompraId ORDER BY referencia.ReferenciaId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Historial,
            (SELECT COUNT_BIG(1) FROM dbo.ReferenciaComprobanteExterno AS referencia WHERE referencia.OrdenCompraId=oc.OrdenCompraId) AS TotalHistorial,
            JSON_QUERY((SELECT evento.EventoId,evento.Accion,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
                FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
                WHERE evento.Entidad=N'ReferenciaComprobanteExterno' AND EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS vinculo WHERE vinculo.OrdenCompraId=oc.OrdenCompraId AND CONVERT(NVARCHAR(100),vinculo.ReferenciaId)=evento.RegistroId)
                ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Auditoria,
            (SELECT COUNT_BIG(1) FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'ReferenciaComprobanteExterno' AND EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS vinculo WHERE vinculo.OrdenCompraId=oc.OrdenCompraId AND CONVERT(NVARCHAR(100),vinculo.ReferenciaId)=evento.RegistroId)) AS TotalAuditoria
        FROM dbo.OrdenCompraCliente AS oc LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1 WHERE oc.OrdenCompraId=@OrdenCompraId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
