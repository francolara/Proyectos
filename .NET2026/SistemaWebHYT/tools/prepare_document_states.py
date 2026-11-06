"""Genera transiciones documentales auditadas; no conecta con SQL."""
from pathlib import Path
import subprocess
import sys
from document_schema import DOCUMENTS
from prepare_followup import procedure

PAIRS={
 'OC':[('REGISTRADA','TRABAJO_COMPLETADO'),('REGISTRADA','CANCELADA'),('TRABAJO_COMPLETADO','CERRADA_OPERATIVAMENTE'),('TRABAJO_COMPLETADO','REGISTRADA'),('TRABAJO_COMPLETADO','CANCELADA'),('CERRADA_OPERATIVAMENTE','TRABAJO_COMPLETADO'),('CANCELADA','REGISTRADA')],
 'GUIA':[(a,b) for a,targets in {'REGISTRADA':['ENVIADA','OBSERVADA','ACEPTADA','ANULADA'],'ENVIADA':['OBSERVADA','ACEPTADA','ANULADA'],'OBSERVADA':['ENVIADA','ACEPTADA','ANULADA'],'ACEPTADA':['OBSERVADA','ANULADA'],'ANULADA':['REGISTRADA']}.items() for b in targets],
 'CONFIRMACION':[('VIGENTE','RETIRADA')],
 'EJECUCION':[('PENDIENTE','PROGRAMADA'),('PENDIENTE','EN_EJECUCION'),('PENDIENTE','CANCELADA'),('PROGRAMADA','PENDIENTE'),('PROGRAMADA','EN_EJECUCION'),('PROGRAMADA','CANCELADA'),('EN_EJECUCION','TERMINADA'),('EN_EJECUCION','CANCELADA'),('TERMINADA','EN_EJECUCION'),('CANCELADA','PENDIENTE')]
}
for entity in DOCUMENTS:
    code=entity['code'];table=entity['table'];pk=entity['pk'];parent=entity['parent']
    state="CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END" if code=='CONFIRMACION' else 'documento.Estado'
    cols=[f['name'] for f in entity['fields']]+(['Vigente'] if code=='CONFIRMACION' else ['Estado'])
    if code=='GUIA':cols+=['FechaAceptacion','RepresentanteCliente']
    old=','.join('documento.'+c for c in cols);new=','.join('actual.'+c for c in cols)
    body=f'''        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'{entity['permission']}.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'{code}_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT={'(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId)' if code=='CONFIRMACION' else '@PadreId'},@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
'''
    if code=='CONFIRMACION':
        body+='''        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió. Vuelve a consultarla.',1;
'''
    else:
        body+='''        IF NOT EXISTS(SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
'''
    body+=f'''        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior={state},@Anterior=(SELECT {old} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.{table} AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.{pk}=@Id AND documento.{parent}=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
'''
    transitions=' OR '.join(f"(@EstadoAnterior=N'{a}' AND @NuevoEstado=N'{b}')" for a,b in PAIRS[code])
    body+=f'''        IF NOT({transitions}) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
'''
    if code=='GUIA':
        body+='''        IF @EstadoAnterior=N'ACEPTADA' AND EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado=N'CERRADA_OPERATIVAMENTE')
            THROW 51433,N'[HYT_VALIDACION] Reabre primero la OC antes de modificar la aceptación de la guía.',1;
        UPDATE dbo.GuiaConformidad SET Estado=@NuevoEstado,FechaAceptacion=CASE WHEN @NuevoEstado=N'ACEPTADA' THEN CONVERT(DATETIME2(3),@FechaEvento) END,
            RepresentanteCliente=CASE WHEN @NuevoEstado=N'ACEPTADA' THEN @Responsable END,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE GuiaId=@Id;
'''
    elif code=='OC':
        body+='''        IF @NuevoEstado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')
            THROW 51433,N'[HYT_VALIDACION] Registra primero la ejecución terminada del trabajo.',1;
        IF @NuevoEstado=N'CERRADA_OPERATIVAMENTE' AND NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId AND guia.Estado=N'ACEPTADA')
            THROW 51433,N'[HYT_VALIDACION] Para cerrar la OC se requiere la guía aceptada.',1;
        IF @NuevoEstado=N'CANCELADA' AND (EXISTS(SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1)
            OR EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId AND guia.Estado<>N'ANULADA')
            OR EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'EN_EJECUCION'))
            THROW 51433,N'[HYT_VALIDACION] Retira la confirmación, anula la guía y detén la ejecución antes de cancelar la OC.',1;
        UPDATE dbo.OrdenCompraCliente SET Estado=@NuevoEstado,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@Id;
'''
    elif code=='CONFIRMACION':
        body+='''        UPDATE dbo.ConfirmacionFacturacion SET Vigente=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@PadreId;
'''
    else:
        body+='''        IF @NuevoEstado=N'PROGRAMADA' AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            THROW 51433,N'[HYT_VALIDACION] Completa fecha programada y responsable del trabajo en la ficha.',1;
        IF @NuevoEstado=N'EN_EJECUCION' AND (NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            OR (@TipoAtencion=N'INVITACION' AND (NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado<>N'CANCELADA')
                OR NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL))))
            THROW 51433,N'[HYT_VALIDACION] Completa responsable; invitación requiere OC y programación antes de iniciar.',1;
        IF @NuevoEstado=N'TERMINADA' AND EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND (ejecucion.FechaInicioReal IS NULL OR CONVERT(DATE,ejecucion.FechaInicioReal)>@FechaEvento))
            THROW 51433,N'[HYT_VALIDACION] La fecha de término debe ser igual o posterior al inicio real.',1;
        IF @EstadoAnterior=N'TERMINADA' AND (EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE'))
            OR EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId AND guia.Estado=N'ACEPTADA'))
            THROW 51433,N'[HYT_VALIDACION] Reabre la OC y revisa la aceptación de la guía antes de reabrir el trabajo.',1;
        UPDATE dbo.Ejecucion SET Estado=@NuevoEstado,
            FechaInicioReal=CASE WHEN @NuevoEstado=N'EN_EJECUCION' THEN COALESCE(FechaInicioReal,CONVERT(DATETIME2(3),@FechaEvento)) ELSE FechaInicioReal END,
            FechaTerminoReal=CASE WHEN @NuevoEstado=N'TERMINADA' THEN CONVERT(DATETIME2(3),@FechaEvento) WHEN @NuevoEstado=N'EN_EJECUCION' THEN NULL ELSE FechaTerminoReal END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE EjecucionId=@Id;
'''
    body+=f'''        SELECT @Nuevo=(SELECT {new} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.{table} AS actual WHERE actual.{pk}=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'{table}',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
'''
    procedure('spu_Seguimiento'+code+'_CambiarEstado','Cambia estado con dependencias, fecha, responsable, auditoría e idempotencia.',
        '    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),\n    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL',body,True)
subprocess.run([sys.executable,str(Path(__file__).with_name('prepare_followup.py'))],check=True)