-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta Orden de Compra - Cliente, contexto, adjuntos e historial paginado.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.PresupuestoId FROM dbo.OrdenCompraCliente AS documento WHERE documento.OrdenCompraId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.OrdenCompraId FROM dbo.OrdenCompraCliente AS documento WHERE documento.PresupuestoId=@PadreId;
        DECLARE @PresupuestoId BIGINT;
        SET @PresupuestoId=@PadreId;
        SELECT (SELECT N'OC' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            p.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.OrdenCompraId, documento.PresupuestoId, documento.RevisionId, documento.NumeroOCCliente, documento.FechaEmisionOC, documento.FechaRecepcionOC, documento.Importe, documento.Observacion, documento.NumeroRegistroInterno, documento.CodigoMoneda, documento.Estado, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,documento.Estado AS EstadoDocumento
                FROM dbo.OrdenCompraCliente AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.OrdenCompraId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.OrdenCompraId=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'OrdenCompraCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'OrdenCompraCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
