"""Genera fuentes SQL estáticas de seguimiento. La aplicación no construye SQL dinámico."""
from pathlib import Path
import json
import re
from document_schema import DOCUMENTS

BASE=Path(__file__).resolve().parents[3]/'Basededatos/Dbsishyt'
def save(path,text):path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
def permission(code):
    return f'''        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'{code}')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
'''
def procedure(name,description,parameters,body,transaction=False):
    text=f'''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   {description}
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.{name}
{parameters}
AS
BEGIN
    SET NOCOUNT ON;
'''
    if transaction:text+='    SET XACT_ABORT ON;\n    DECLARE @Propia BIT=0;\n'
    text+='    BEGIN TRY\n'+body+'    END TRY\n    BEGIN CATCH\n'
    if transaction:text+='        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;\n'
    if '#Pagina' in body:text+="        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;\n"
    text+='''        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
'''
    if name.endswith("_Listar"):
        text=text.replace("USE [SistemaWebHYT];","-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.\nUSE [SistemaWebHYT];",1)
    save(BASE/'StoreProcedure'/f'{name}.sql',text)

def state(e,alias):return f"CASE WHEN {alias}.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END" if e['code']=='CONFIRMACION' else alias+'.Estado'

for e in DOCUMENTS:
    table,pk,code,parent=e['table'],e['pk'],e['code'],e['parent']
    body=permission(e['permission']+'.EDITAR')+'''        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
'''
    for f in e['fields']:
        name,typ=f['name'],f['sql'];raw=f"JSON_VALUE(@Datos,N'$.{name}')"
        length=re.search(r'NVARCHAR\((\d+)\)',typ)
        if length:
            body+=f"        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'{name}' AND LEN(campo.value)>{length.group(1)}) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: {name}.',1;\n"
            expr=f"NULLIF(LTRIM(RTRIM({raw})),N'')"
        else:
            expr=f"TRY_CONVERT({typ},NULLIF({raw},N''))"
            body+=f"        IF NULLIF({raw},N'') IS NOT NULL AND {expr} IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: {name}.',1;\n"
            if f['kind']=='Date':body+=f"        IF LEN({raw})>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: {name}.',1;\n"
        body+=f'        DECLARE @{name} {typ}={expr};\n'
        if f['required']:body+=f"        IF @{name} IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: {name}.',1;\n"
        if f['kind']=='Money':body+=f"        IF @{name}<0 OR ({raw} IS NOT NULL AND (PATINDEX(N'%[^0-9.]%',{raw})>0 OR (CHARINDEX(N'.',{raw})>0 AND LEN({raw})-CHARINDEX(N'.',{raw})>2))) THROW 51408,N'[HYT_VALIDACION] Importe no negativo con hasta dos decimales.',1;\n"
    body+=f'''        DECLARE @TipoOperacion NVARCHAR(40)=N'{code}_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'{code}_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'{e['state']}',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
'''
    if code=='CONFIRMACION':body+='''        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        IF @PresupuestoId IS NULL THROW 51410,N'[HYT_VALIDACION] La OC del cliente no existe.',1;
'''
    else:body+='        SET @PresupuestoId=@PadreId;\n'
    body+='''        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
'''
    if code=='CONFIRMACION':body+='''        IF NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre AND oc.Estado<>N'CANCELADA')
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió o está cancelada. Vuelve a consultarla.',1;
'''
    else:body+='''        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
'''
    body+='''        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de modificar documentos del expediente.',1;
'''
    oldfields=[f['name'] for f in e['fields']]
    if code=='OC':oldfields+=['NumeroRegistroInterno','CodigoMoneda']
    if code=='GUIA':oldfields+=['FechaAceptacion','RepresentanteCliente']
    if code=='CONFIRMACION':oldfields+=['Vigente']
    else:oldfields+=['Estado']
    selected=', '.join('documento.'+name for name in oldfields)
    body+=f'''        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT {selected} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado={state(e,'documento')}
            FROM dbo.{table} AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.{pk}=@Id AND documento.{parent}=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
'''
    condition=' AND documento.Vigente=1' if code=='CONFIRMACION' else ''
    body+=f"        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.{table} AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.{parent}=@PadreId{condition}) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;\n"
    if code=='OC':body+='''        DECLARE @EstadoRevision NVARCHAR(20),@TotalRevision DECIMAL(19,2);
        SELECT @EstadoRevision=r.Estado,@TotalRevision=r.Total,@Moneda=r.CodigoMoneda FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId AND r.RevisionId=@RevisionId;
        IF @EstadoRevision IS NULL OR @TotalRevision IS NULL OR (@TipoAtencion=N'INVITACION' AND @EstadoRevision<>N'APROBADO') OR (@TipoAtencion=N'EMERGENCIA' AND @EstadoRevision NOT IN (N'BORRADOR',N'ENVIADO',N'APROBADO'))
            THROW 51414,N'[HYT_VALIDACION] Invitación requiere revisión aprobada. Emergencia admite revisión completa sin aprobación previa.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la OC antes de editarla.',1;
        IF @Id IS NOT NULL AND (@Estado=N'CERRADA_OPERATIVAMENTE' OR EXISTS (SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1))
            AND (@RevisionId<>TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.RevisionId')) OR @NumeroOCCliente<>JSON_VALUE(@Anterior,N'$.NumeroOCCliente')
                OR @FechaEmisionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaEmisionOC'),23) OR @FechaRecepcionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaRecepcionOC'),23)
                OR COALESCE(@Importe,-1)<>COALESCE(TRY_CONVERT(DECIMAL(19,2),JSON_VALUE(@Anterior,N'$.Importe')),-1))
            THROW 51415,N'[HYT_VALIDACION] Reabre la OC o retira su confirmación antes de corregir sus datos principales.',1;
'''
    elif code=='GUIA':body+='''        IF @TipoAtencion=N'INVITACION' AND NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.PresupuestoRevision AS r ON r.RevisionId=oc.RevisionId AND r.Estado=N'APROBADO' WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado<>N'CANCELADA')
            THROW 51416,N'[HYT_VALIDACION] Invitación requiere una OC del cliente vinculada a revisión aprobada.',1;
        IF @Estado=N'ANULADA' THROW 51415,N'[HYT_VALIDACION] La guía anulada no permite modificaciones.',1;
        IF @Id IS NOT NULL AND @Estado=N'ACEPTADA' AND (@NumeroActa<>JSON_VALUE(@Anterior,N'$.NumeroActa') OR @FechaActa<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaActa'),23))
            THROW 51415,N'[HYT_VALIDACION] Registra observación de la guía antes de corregir número o fecha aceptados.',1;
'''
    elif code=='CONFIRMACION':body+="        IF @Estado=N'RETIRADA' THROW 51415,N'[HYT_VALIDACION] La confirmación retirada se conserva como historial. Registra una nueva.',1;\n"
    else:body+='''        IF @FechaTerminoReal IS NOT NULL AND (@FechaInicioReal IS NULL OR @FechaTerminoReal<@FechaInicioReal) THROW 51417,N'[HYT_VALIDACION] Revisa inicio y término reales.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la ejecución antes de editarla.',1;
        IF @Estado=N'PROGRAMADA' AND @FechaProgramada IS NULL THROW 51417,N'[HYT_VALIDACION] La ejecución programada necesita fecha.',1;
        IF @Estado IN (N'EN_EJECUCION',N'TERMINADA') AND (@FechaInicioReal IS NULL OR @Responsable IS NULL OR (@Estado=N'TERMINADA' AND @FechaTerminoReal IS NULL))
            THROW 51417,N'[HYT_VALIDACION] Conserva responsable y fechas reales según el estado.',1;
'''
    names=[f['name'] for f in e['fields']]+[parent,'UsuarioRegistro'];values=['@'+f['name'] for f in e['fields']]+['@PadreId','@UsuarioId']
    if code=='OC':
        names+=['NumeroRegistroInterno','CodigoMoneda'];values+=['@NumeroInterno','@Moneda']
        body+='''        DECLARE @NumeroInterno NVARCHAR(40);
        IF @Id IS NULL EXEC dbo.spu_ObtenerNumeroInterno @TipoRegistro=N'OC_CLIENTE',@NumeroInterno=@NumeroInterno OUTPUT;
'''
    body+=f'''        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.{table} ({','.join(names)}) VALUES ({','.join(values)});
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.{table} SET {','.join(f['name']+'=@'+f['name'] for f in e['fields'])}{',CodigoMoneda=@Moneda' if code=='OC' else ''},FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE {pk}=@Id;
'''
    if e['filetable']:
        body+='''        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
'''
        body+=f'''        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.{e['filetable']} AS vinculo WHERE vinculo.{e['filekey']}=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.{e['filetable']} AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.{e['filekey']}=@Id
        WHERE archivo.Retirado=0 AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE TRY_CONVERT(BIGINT,entrada.value)=archivo.ArchivoId);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoAnterior,EstadoNuevo,UsuarioId,Motivo,CambiosJson)
        SELECT @CorrelacionId,N'ArchivoAdjunto',CONVERT(NVARCHAR(100),retirado.ArchivoId),N'RETIRAR',N'VIGENTE',N'RETIRADO',@UsuarioId,@Motivo,
            (SELECT retirado.NombreOriginal AS NombreOriginal FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM @ArchivosRetirados AS retirado;
'''
        body+='''        IF @ArchivoJson IS NOT NULL
        BEGIN
            DECLARE @Nombre NVARCHAR(255)=JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'),@Clave NVARCHAR(500)=JSON_VALUE(@ArchivoJson,N'$.ClaveAlmacenamiento'),@TipoContenido NVARCHAR(150)=JSON_VALUE(@ArchivoJson,N'$.TipoContenido'),
                @Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@ArchivoJson,N'$.TamanoBytes')),@SHA CHAR(64)=JSON_VALUE(@ArchivoJson,N'$.SHA256'),@Extension NVARCHAR(5);
            SET @Extension=CASE @TipoContenido WHEN N'application/pdf' THEN N'.pdf' WHEN N'image/png' THEN N'.png' WHEN N'image/jpeg' THEN N'.jpg' END;
            IF NULLIF(@Nombre,N'') IS NULL OR @Tamano IS NULL OR @Tamano<=0 OR @Tamano>15728640 OR @Extension IS NULL OR LEN(JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'))>255
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR @Clave<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
'''
        if code=='GUIA':body+="            IF @TipoContenido<>N'application/pdf' THROW 51419,N'[HYT_VALIDACION] La guía solo admite escaneado PDF.',1;\n"
        body+=f'''            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.{e['filetable']} ({e['filekey']},ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
'''
    else:body+="        IF JSON_QUERY(@Datos,N'$.Adjunto') IS NOT NULL THROW 51419,N'[HYT_VALIDACION] Esta ficha no admite adjuntos.',1;\n"
    body+=f'''        SELECT @Nuevo=(SELECT {selected.replace('documento.','actual.')} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.{table} AS actual WHERE actual.{pk}=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'{table}',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
'''
    if code=='CONFIRMACION':body+='        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@PadreId;\n'
    body+='''        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
'''
    procedure('spu_Seguimiento'+code+'_Guardar',f'Guarda {e["title"]} con idempotencia, concurrencia, archivos y auditoría.',
        '    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),\n    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL',body,True)

for e in DOCUMENTS:
    table,pk,code,parent=e['table'],e['pk'],e['code'],e['parent']
    body=permission(e['permission']+'.CONSULTAR')+'''        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
'''
    body+=f'''        IF @Id IS NOT NULL SELECT @PadreReal=documento.{parent} FROM dbo.{table} AS documento WHERE documento.{pk}=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.{pk} FROM dbo.{table} AS documento WHERE documento.{parent}=@PadreId{' AND documento.Vigente=1' if code=='CONFIRMACION' else ''};
        DECLARE @PresupuestoId BIGINT;
'''
    if code=='CONFIRMACION':body+='        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;\n'
    else:body+='        SET @PresupuestoId=@PadreId;\n'
    fields=[pk,parent]+[f['name'] for f in e['fields']]
    if code=='OC':fields+=['NumeroRegistroInterno','CodigoMoneda']
    if code=='GUIA':fields+=['FechaAceptacion','RepresentanteCliente']
    if code=='CONFIRMACION':fields+=['Vigente']
    else:fields+=['Estado']
    fields+=['FechaRegistro','FechaActualizacion','VersionFila']
    docfields=', '.join('documento.'+name for name in fields)
    body+=f'''        SELECT (SELECT N'{code}' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroInterno AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            {'oc.VersionFila' if code=='CONFIRMACION' else 'p.VersionFila'} AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT {docfields},creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,{state(e,'documento')} AS EstadoDocumento
                FROM dbo.{table} AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.{pk}=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
'''
    if e['filetable']:
        body+=f'''            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.{e['filetable']} AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.{e['filekey']}=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
'''
    else:body+="            JSON_QUERY(N'[]') AS Adjuntos\n"
    body+=f'''        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'{table}' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'{table}' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
'''
    procedure('spu_Seguimiento'+code+'_Obtener',f'Consulta {e["title"]}, contexto, adjuntos e historial paginado.',
        '    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1',body)

    root='oc.PresupuestoId' if code=='CONFIRMACION' else 'documento.PresupuestoId'
    joins=' INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=documento.OrdenCompraId' if code=='CONFIRMACION' else ''
    if code=='OC':revision=' INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=documento.RevisionId'
    elif code=='CONFIRMACION':revision=' INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId'
    else:revision=' OUTER APPLY (SELECT TOP (1) r.ClienteNombre,r.TiendaNombre,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente'
    datefield={'OC':'FechaEmisionOC','GUIA':'FechaActa','CONFIRMACION':'FechaConfirmacion','EJECUCION':'FechaProgramada'}[code]
    number=f'CONVERT(NVARCHAR(100),documento.{e["number"]})'
    external='documento.NumeroOCCliente' if code=='OC' else number
    if code=='OC':summary=' LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=documento.OrdenCompraId AND confirmacion.Vigente=1'
    elif code=='CONFIRMACION':summary=' LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.ConfirmacionId=documento.ConfirmacionId'
    else:summary=' LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1'
    files=f"(SELECT COUNT_BIG(1) FROM dbo.{e['filetable']} AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.{e['filekey']}=documento.{pk})" if e['filetable'] else 'CONVERT(BIGINT,0)'
    firstfile=f"(SELECT MIN(archivo.ArchivoId) FROM dbo.{e['filetable']} AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.{e['filekey']}=documento.{pk})" if e['filetable'] else 'CONVERT(BIGINT,NULL)'
    body=permission(e['permission']+'.CONSULTAR')+'''        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
'''
    invoice="(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.FechaRegistro,autor.Nombres AS UsuarioRegistro FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.AspNetUsers AS autor ON autor.Id=referencia.UsuarioRegistro WHERE referencia.OrdenCompraId=documento.OrdenCompraId AND referencia.Vigente=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permisoFactura ON permisoFactura.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permisoFactura.CodigoPermiso=N'FACTURADO.CONSULTAR') FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)" if code=='OC' else 'CONVERT(NVARCHAR(MAX),NULL)'
    body+=f'''        SELECT documento.{pk} AS Id,documento.{parent} AS PadreId,p.PresupuestoId,p.NumeroInterno AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            {number} AS NumeroInterno,{external} AS NumeroDocumento,{state(e,'documento')} AS Estado,documento.{datefield} AS FechaDocumento,
            fuente.CodigoMoneda,{'documento.Importe' if code=='OC' else 'CONVERT(DECIMAL(19,2),NULL)'} AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),{'oc.VersionFila' if code=='CONFIRMACION' else 'p.VersionFila'}) AS VersionPadre,{files} AS Adjuntos,{firstfile} AS PrimerArchivoId,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,{invoice} AS Facturado
        INTO #Pagina FROM dbo.{table} AS documento{joins}
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId={root}{revision}{summary}
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR {state(e,'documento')}=@Estado)
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR {number} LIKE @Patron ESCAPE N'~' OR {external} LIKE @Patron ESCAPE N'~' OR p.NumeroInterno LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
'''
    procedure('spu_Seguimiento'+code+'_Listar',f'Lista {e["title"]} por fecha/búsqueda, adjuntos y confirmación; veinte por página.',
        '    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL',body)

sources=['Tablas/DocumentoOperacion.sql','Script/016_SeguimientoDocumental.sql']+sorted('StoreProcedure/'+path.name for path in (BASE/'StoreProcedure').glob('spu_Seguimiento*.sql'))
grant='Script/017_PermisosEtapa6.sql'
text='''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos individuales de seguimiento documental, etapa 6.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
'''
text+=''.join(f'GRANT EXECUTE ON OBJECT::dbo.{Path(path).stem} TO hyt_app_executor;\n' for path in sources if path.startswith('StoreProcedure/'))
save(BASE/grant,text+'GO\n');sources.append(grant)
save(BASE/'manifest-etapa6.json',json.dumps({'database':'SistemaWebHYT','requires':'etapa5','scripts':sources},indent=2)+'\n')
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'));prior=[]
for stage in range(1,6):prior+=json.loads((BASE/f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))['scripts']
manifest['scripts']=prior+sources
for later in range(7,10):
    if (BASE/f'manifest-etapa{later}.json').exists():manifest['scripts']+=json.loads((BASE/f'manifest-etapa{later}.json').read_text(encoding='utf-8'))['scripts']
save(BASE/'manifest.json',json.dumps(manifest,indent=2)+'\n')
print(f'Seguimiento: {len(sources)} fuentes registradas; sin conexión SQL desde el generador.')
