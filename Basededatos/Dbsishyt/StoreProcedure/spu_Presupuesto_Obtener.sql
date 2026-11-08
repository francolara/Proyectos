-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Recupera revisión y partidas históricas, revisiones y auditoría del expediente paginada y filtrada por permisos.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible por año/mes de registro en Lima y correlativo mensual. Centro SAP FICO desde tienda cuando falta en partidas anteriores. Contactos activos de la tienda para impresión anterior.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT,@RevisionId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        IF @RevisionId IS NULL SELECT TOP (1) @RevisionId=r.RevisionId FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@Id ORDER BY r.NumeroRevision DESC;
        IF @RevisionId IS NULL BEGIN SELECT CONVERT(NVARCHAR(MAX),NULL) AS Datos;RETURN;END;
        DECLARE @Referencias TABLE(Entidad NVARCHAR(128) NOT NULL,RegistroId NVARCHAR(100) NOT NULL,PRIMARY KEY(Entidad,RegistroId));
        DECLARE @Permisos TABLE(Codigo NVARCHAR(100) NOT NULL PRIMARY KEY);
        INSERT INTO @Permisos(Codigo) SELECT DISTINCT permiso.CodigoPermiso FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        INSERT INTO @Referencias(Entidad,RegistroId) VALUES(N'Presupuesto',CONVERT(NVARCHAR(100),@Id)),(N'PresupuestoRevision',CONVERT(NVARCHAR(100),@RevisionId));
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'OrdenCompraCliente',CONVERT(NVARCHAR(100),oc.OrdenCompraId) FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'OC_CLIENTE.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'GuiaConformidad',CONVERT(NVARCHAR(100),guia.GuiaId) FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'GUIAS.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'Ejecucion',CONVERT(NVARCHAR(100),ejecucion.EjecucionId) FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'EJECUCION.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'ConfirmacionFacturacion',CONVERT(NVARCHAR(100),confirmacion.ConfirmacionId) FROM dbo.ConfirmacionFacturacion AS confirmacion INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=confirmacion.OrdenCompraId WHERE oc.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'CONFIRMACIONES.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId)
        SELECT N'ArchivoAdjunto',CONVERT(NVARCHAR(100),archivo.ArchivoId) FROM dbo.ArchivoAdjunto AS archivo
        WHERE EXISTS(SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'OrdenCompraCliente' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.OrdenCompraId) WHERE vinculo.ArchivoId=archivo.ArchivoId)
           OR EXISTS(SELECT 1 FROM dbo.GuiaArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'GuiaConformidad' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.GuiaId) WHERE vinculo.ArchivoId=archivo.ArchivoId)
           OR EXISTS(SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'ConfirmacionFacturacion' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.ConfirmacionId) WHERE vinculo.ArchivoId=archivo.ArchivoId);
        SELECT (SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.ClienteId,p.TiendaId,p.TipoAtencion,p.Descripcion,p.FechaSolicitud,p.FechaRegistro,p.FechaActualizacion,p.VersionFila,
            creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,contactosTienda.Nombres AS ContactosTienda,
            JSON_QUERY((SELECT revision.RevisionId, revision.PresupuestoId, revision.NumeroRevision, revision.FechaEmision, revision.NumeroCotizacionCliente, revision.ClienteNombre, revision.ClienteDocumento, revision.ClienteDireccion, revision.TiendaNombre, revision.TiendaDireccion, revision.FechaEjecucion, revision.Trimestre, revision.VigenciaTexto, revision.CodigoSAPProveedor, revision.ContratoMarco, revision.Especialidad, revision.Coordinador, revision.Encargado, revision.Consideraciones, revision.Firmante, revision.CargoFirmante, revision.CodigoMoneda, revision.CodigoTipoIGV, revision.FechaDesdeTasaIGV, revision.TasaIGVAplicada, revision.Estado, revision.Subtotal, revision.Impuesto, revision.Total, revision.FechaRegistro, revision.UsuarioRegistro, revision.FechaActualizacion, revision.UsuarioActualizacion, revision.VersionFila, revision.ZonaTexto, revision.EmpresaSnapshotJson, revision.PoliticaCalculo, revision.FechaAprobacion, revision.DescripcionTrabajo, revision.SimboloMoneda, revision.TiendaCodigo FROM dbo.PresupuestoRevision AS revision WHERE revision.RevisionId=@RevisionId AND revision.PresupuestoId=p.PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Revision,
            JSON_QUERY((SELECT partida.PartidaId, partida.NumeroLinea, partida.Grupo, partida.MaterialId, partida.PartidaTarifarioId, partida.Posicion, partida.CodigoInterno, partida.CodigoMaterialSAP, COALESCE(NULLIF(LTRIM(RTRIM(partida.CentroSAPFICO)),N''),tienda.CentroSAPFICO) AS CentroSAPFICO, partida.NumeroNecesidad, partida.TipoServicio, partida.Descripcion, partida.Unidad, partida.Cantidad, partida.PrecioOrigen, partida.CodigoMonedaOrigen, partida.CodigoTipoIGVOrigen, partida.PrecioUnitario, partida.BaseImponible, partida.Impuesto, partida.Total, partida.ZonasTexto, partida.OrigenPrecio, partida.MotivoPrecioManual FROM dbo.PresupuestoPartida AS partida WHERE partida.RevisionId=@RevisionId ORDER BY partida.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Partidas,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.Estado,r.FechaEmision,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia
        FROM dbo.Presupuesto AS p INNER JOIN dbo.Tienda AS tienda ON tienda.TiendaId=p.TiendaId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        OUTER APPLY (
            SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),NULLIF(LTRIM(RTRIM(contacto.Nombre)),N'')),N' / ') WITHIN GROUP(ORDER BY contacto.ContactoId) AS Nombres
            FROM dbo.ContactoCliente AS contacto
            WHERE contacto.ClienteId=p.ClienteId AND contacto.TiendaId=p.TiendaId AND contacto.Activo=1
        ) AS contactosTienda
        WHERE p.PresupuestoId=@Id AND EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@Id)
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson,evento.Entidad
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
