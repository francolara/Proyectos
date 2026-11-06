"""Fuentes SQL de marcado externo; no emite comprobantes ni conecta con SQL."""
from pathlib import Path
import json
BASE=Path(__file__).resolve().parents[3]/'Basededatos/Dbsishyt'
def save(path,text):path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
def procedure(name,parameters,body,write=False):
    text=f'''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Marcado Facturado externo con permisos y trazabilidad.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.{name}
{parameters}
AS
BEGIN
    SET NOCOUNT ON;
'''
    if write:text+='    SET XACT_ABORT ON;\n    DECLARE @Propia BIT=0;\n'
    text+='    BEGIN TRY\n'+body+'    END TRY\n    BEGIN CATCH\n'
    if write:text+='        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;\n'
    text+='''        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
'''
    source='StoreProcedure/'+name+'.sql';save(BASE/source,text);return source
def permission(action):return f'''        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'FACTURADO.{action}')
            THROW 51500,N'[HYT_PERMISO] No tiene permiso para esta referencia de facturación.',1;
'''
body=permission('EDITAR')+'''        IF @@TRANCOUNT<>0 THROW 51501,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Serie=NULLIF(LTRIM(RTRIM(@Serie)),N'');SET @NumeroDocumento=NULLIF(LTRIM(RTRIM(@NumeroDocumento)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @OrdenCompraId IS NULL OR @OrdenCompraId<=0 OR @VersionOC IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @Retirar IS NULL
            THROW 51502,N'[HYT_VALIDACION] Vuelve a abrir la OC.',1;
        IF @Retirar=0 AND (@FechaDocumento IS NULL OR @CodigoTipoDocumento IS NULL OR @CodigoTipoDocumento NOT IN(N'01',N'03') OR @Serie IS NULL OR @NumeroDocumento IS NULL)
            THROW 51502,N'[HYT_VALIDACION] Completa fecha, tipo 01/03, serie y número del documento externo.',1;
        IF (@ReferenciaId IS NOT NULL OR @Retirar=1) AND @Motivo IS NULL THROW 51502,N'[HYT_VALIDACION] Indica el motivo de la corrección o retiro.',1;
        IF @Retirar=1 AND @ReferenciaId IS NULL THROW 51502,N'[HYT_VALIDACION] Selecciona la referencia a retirar.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @OrdenCompraId AS OrdenCompraId,@VersionOC AS VersionOC,@ReferenciaId AS ReferenciaId,@VersionReferencia AS VersionReferencia,@FechaDocumento AS FechaDocumento,@CodigoTipoDocumento AS Tipo,@Serie AS Serie,@NumeroDocumento AS Numero,@Retirar AS Retirar,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Resultado BIGINT,@Actor NVARCHAR(128),@HashPrevio CHAR(64),@TipoPrevio NVARCHAR(40),@PadrePrevio BIGINT;
        BEGIN TRANSACTION;SET @Propia=1;
        SELECT @Resultado=operacion.RegistroId,@Actor=operacion.UsuarioId,@HashPrevio=operacion.Huella,@TipoPrevio=operacion.TipoOperacion,@PadrePrevio=operacion.PadreId FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Actor<>@UsuarioId OR @HashPrevio<>@Huella OR @TipoPrevio<>N'FACTURADO_GUARDAR' OR @PadrePrevio<>@OrdenCompraId THROW 51503,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS ReferenciaId;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@OrdenCompraId),@Bloqueo BIGINT;
        SELECT @Bloqueo=p.PresupuestoId FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @Bloqueo IS NULL THROW 51504,N'[HYT_VALIDACION] La OC no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@OrdenCompraId AND oc.VersionFila=@VersionOC AND oc.Estado<>N'CANCELADA')
            THROW 51505,N'[HYT_CONCURRENCIA] La OC cambió o está cancelada. Vuelve a consultarla.',1;
        DECLARE @ConfirmacionId BIGINT,@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@EstadoAnterior NVARCHAR(30);
        SELECT @ConfirmacionId=confirmacion.ConfirmacionId FROM dbo.ConfirmacionFacturacion AS confirmacion WITH(UPDLOCK,HOLDLOCK) WHERE confirmacion.OrdenCompraId=@OrdenCompraId AND confirmacion.Vigente=1;
        IF @Retirar=0 AND @ConfirmacionId IS NULL THROW 51506,N'[HYT_VALIDACION] Registra primero la confirmación del cliente de esta OC.',1;
        IF @Retirar=0 AND (NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId AND guia.Estado=N'ACEPTADA') OR NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA'))
            THROW 51506,N'[HYT_VALIDACION] Completa el trabajo y registra la guía aceptada antes de marcar Facturado.',1;
        IF @ReferenciaId IS NULL AND EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS referencia WITH(UPDLOCK,HOLDLOCK) WHERE referencia.OrdenCompraId=@OrdenCompraId AND referencia.Vigente=1)
            THROW 51507,N'[HYT_VALIDACION] La OC ya está marcada. Abre su referencia para corregirla.',1;
        IF @ReferenciaId IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.Vigente,referencia.ConfirmacionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ReferenciaComprobanteExterno AS referencia WITH(UPDLOCK,HOLDLOCK) WHERE referencia.ReferenciaId=@ReferenciaId AND referencia.OrdenCompraId=@OrdenCompraId AND referencia.VersionFila=@VersionReferencia AND referencia.Vigente=1;
            IF @Anterior IS NULL THROW 51505,N'[HYT_CONCURRENCIA] La referencia cambió. Vuelve a consultarla.',1;
            SET @EstadoAnterior=N'FACTURADO';
        END;
        IF @ReferenciaId IS NULL
        BEGIN
            INSERT INTO dbo.ReferenciaComprobanteExterno(OrdenCompraId,ConfirmacionId,FechaDocumento,CodigoTipoDocumento,Serie,NumeroDocumento,UsuarioRegistro) VALUES(@OrdenCompraId,@ConfirmacionId,@FechaDocumento,@CodigoTipoDocumento,@Serie,@NumeroDocumento,@UsuarioId);
            SET @ReferenciaId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE IF @Retirar=1
            UPDATE dbo.ReferenciaComprobanteExterno SET Vigente=0,MotivoRetiro=@Motivo,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ReferenciaId=@ReferenciaId;
        ELSE
            UPDATE dbo.ReferenciaComprobanteExterno SET ConfirmacionId=@ConfirmacionId,FechaDocumento=@FechaDocumento,CodigoTipoDocumento=@CodigoTipoDocumento,Serie=@Serie,NumeroDocumento=@NumeroDocumento,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ReferenciaId=@ReferenciaId;
        SELECT @Nuevo=(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.Vigente,referencia.ConfirmacionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ReferenciaComprobanteExterno AS referencia WHERE referencia.ReferenciaId=@ReferenciaId;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@ReferenciaId),@Accion NVARCHAR(40)=CASE WHEN @Retirar=1 THEN N'RETIRAR_FACTURADO' WHEN @Anterior IS NULL THEN N'MARCAR_FACTURADO' ELSE N'CORREGIR_FACTURADO' END,
            @EstadoNuevo NVARCHAR(30)=CASE WHEN @Retirar=1 THEN N'RETIRADO' ELSE N'FACTURADO' END,@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ReferenciaComprobanteExterno',@Registro,@Accion,@EstadoAnterior,@EstadoNuevo,@FechaDocumento,@Motivo,@Detalle;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@OrdenCompraId;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,N'FACTURADO_GUARDAR',@OrdenCompraId,@ReferenciaId,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @ReferenciaId AS ReferenciaId;
'''
sources=[procedure('spu_Facturado_Guardar','    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@OrdenCompraId BIGINT,@VersionOC BINARY(8),\n    @FechaDocumento DATE=NULL,@CodigoTipoDocumento NVARCHAR(2)=NULL,@Serie NVARCHAR(30)=NULL,@NumeroDocumento NVARCHAR(50)=NULL,\n    @ReferenciaId BIGINT=NULL,@VersionReferencia BINARY(8)=NULL,@Retirar BIT=0,@Motivo NVARCHAR(1000)=NULL',body,True)]
body=permission('CONSULTAR')+'''        IF @Pagina<1 OR @Pagina>100000000 THROW 51502,N'[HYT_VALIDACION] Página no válida.',1;
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
'''
sources.append(procedure('spu_Facturado_Obtener','    @UsuarioId NVARCHAR(128),@OrdenCompraId BIGINT,@Pagina INT=1',body))
grant='Script/018_PermisosEtapa7.sql'
save(BASE/grant,'-- =============================================\n-- Author:        FRANCO LARA\n-- Create date:   03/11/2026\n-- Description:   Permisos limitados del marcado externo.\n-- =============================================\nUSE [SistemaWebHYT];\nGO\n'+''.join(f'GRANT EXECUTE ON OBJECT::dbo.{Path(p).stem} TO hyt_app_executor;\n' for p in sources)+'GO\n')
sources.append(grant)
save(BASE/'manifest-etapa7.json',json.dumps({'database':'SistemaWebHYT','requires':'etapa6','scripts':sources},indent=2)+'\n')
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))
manifest['scripts']=[p for n in range(1,10) if (BASE/f'manifest-etapa{n}.json').exists() for p in json.loads((BASE/f'manifest-etapa{n}.json').read_text(encoding='utf-8'))['scripts']]
save(BASE/'manifest.json',json.dumps(manifest,indent=2)+'\n')
print('Etapa 7: tres fuentes preparadas, sin ejecución por el generador.')