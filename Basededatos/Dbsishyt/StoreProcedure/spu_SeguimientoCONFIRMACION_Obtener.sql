-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta Confirmación del cliente, contexto, adjuntos e historial paginado.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.OrdenCompraId FROM dbo.ConfirmacionFacturacion AS documento WHERE documento.ConfirmacionId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.ConfirmacionId FROM dbo.ConfirmacionFacturacion AS documento WHERE documento.OrdenCompraId=@PadreId AND documento.Vigente=1;
        DECLARE @PresupuestoId BIGINT;
        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        SELECT (SELECT N'CONFIRMACION' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            oc.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.ConfirmacionId, documento.OrdenCompraId, documento.FechaConfirmacion, documento.NumeroConfirmacion, documento.RepresentanteCliente, documento.Observacion, documento.Vigente, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END AS EstadoDocumento
                FROM dbo.ConfirmacionFacturacion AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.ConfirmacionId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.ConfirmacionId=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ConfirmacionFacturacion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'ConfirmacionFacturacion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
