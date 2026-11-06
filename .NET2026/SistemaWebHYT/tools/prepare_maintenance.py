"""Genera SP estáticos y definiciones C# de etapas 3/4. No ejecuta SQL."""
from pathlib import Path
import json
import re
from maintenance_schema import ENTITIES, BY_KEY
from client_links_sql import SNAPSHOT as LINKS_SNAPSHOT, VALIDATE as LINKS_VALIDATE, SAVE as LINKS_SAVE

ROOT = Path(__file__).resolve().parents[3]
PROJECT = Path(__file__).resolve().parents[1]
BASE = ROOT / 'Basededatos/Dbsishyt'

def save(path,text):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))

def literal(value): return 'null' if value is None else json.dumps(value,ensure_ascii=False)
definitions=[]
for e in ENTITIES:
    fields=[]
    for field in e['fields']:
        length=re.search(r'NVARCHAR\((\d+)\)',field['sql'])
        fields.append('new('+', '.join([literal(field['name']),literal(field['label']),literal(field['kind']),str(field['required']).lower(),str(int(length.group(1)) if length else 0),literal(field['lookup']),literal(field['default']),literal(field['catalog'])])+')')
    definitions.append('        new('+', '.join([literal(e['key']),literal(e['title']),literal(e['table']),literal(e['permission']), '[\n            '+',\n            '.join(fields)+'\n        ]',literal(e['parent']),str(bool(e['zones'])).lower()])+')')
save(PROJECT/'GestionServicios.Application/Maintenance/Definitions.cs','''namespace GestionServicios.Application.Maintenance;

// Firma: FRANCO LARA - 05/11/2026 | Códigos automáticos, abreviaturas, posición en tarifarios y número de documento con tipo oculto, Código SAP, enlaces por cliente y descripción de material de tres líneas.
// Lista cerrada compilada; nunca se aceptan nombres de tabla/SP desde formularios.
public static class MaintenanceDefinitions
{
    public static IReadOnlyList<EntityDefinition> All { get; } =
    [
'''+',\n'.join(definitions)+'''
    ];
    public static EntityDefinition? Find(string key) => All.FirstOrDefault(item => item.Key == key);
}
''')

def permission(e,action):
    return f'''        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'{e['permission']}.{action}'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
'''

scripts=[]
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
    if transaction:text+='    SET XACT_ABORT ON;\n    DECLARE @Propia BIT = 0;\n'
    text+='    BEGIN TRY\n'+body+'    END TRY\n    BEGIN CATCH\n'
    if transaction:text+='        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;\n'
    if name.endswith('_Listar'):text+="        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;\n"
    text+='''        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
'''
    path='StoreProcedure/'+name+'.sql'
    if (e.get('auto_prefix') and name.endswith('_Guardar')) or name in ('spu_MantTienda_Obtener','spu_MantTienda_Listar','spu_MantContactoCliente_Listar','spu_MantTarifario_Listar','spu_MantTarifarioVersion_Listar','spu_MantTarifarioPartida_Listar','spu_MantTarifarioPartida_Obtener','spu_MantCliente_Obtener','spu_MantCliente_Listar','spu_MantServicioMaterial_Obtener','spu_MantServicioMaterial_Listar','spu_MantUnidadMedida_Listar','spu_MantUnidadMedida_Guardar','spu_MantProveedorCliente_Listar'):
        text=text.replace('USE [SistemaWebHYT];','-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.\nUSE [SistemaWebHYT];',1)
    if name in ('spu_MantTienda_Listar','spu_MantContactoCliente_Listar'):text=text.replace('Códigos automáticos inalterables y retiro de posición del maestro de materiales.','Listado y búsqueda por código/nombre del cliente, con veinte filas por página.')
    if name=='spu_MantCliente_Obtener':text=text.replace('Códigos automáticos inalterables y retiro de posición del maestro de materiales.','Consulta identidad, Código SAP y enlaces de documentos ordenados con historial del cliente.')
    if name=='spu_MantCliente_Guardar':text=text.replace('Códigos automáticos inalterables y retiro de posición del maestro de materiales.','Código automático, identidad según persona y número obligatorio; tipo de documento oculto sin validación, conservando el histórico y Código SAP/enlaces múltiples auditados por cliente.')
    if name=='spu_MantUnidadMedida_Guardar':text=text.replace('        RAISERROR (@ErrorMessage',"        IF ERROR_NUMBER() IN(2601,2627) SET @ErrorMessage=N'[HYT_VALIDACION] Ya existe una unidad con esa abreviatura.';\n        RAISERROR (@ErrorMessage")
    save(BASE/path,text);scripts.append(path)

def joins(e):
    if e['key']=='versiones-tarifario': return ' INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId = m.TarifarioId'
    if e['key']=='partidas-tarifario': return ' INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId = m.MaterialId INNER JOIN dbo.TarifarioVersion AS padre ON padre.VersionId = m.VersionId'
    return ''

def label(target,alias):
    if target['key']=='versiones-tarifario':return f"CONCAT(N'Versión ', {alias}.NumeroVersion, N' / Tarifario ', {alias}.TarifarioId)"
    if target['key']=='partidas-tarifario':return f"CONCAT(N'Partida ', {alias}.PartidaId)"
    return target['description'].replace('m.',alias+'.')

for e in ENTITIES:
    table,pk=e['table'],e['pk']
    code,desc,extra=e['code'],e['description'],e['extra']
    if e['key'] in ('clientes','tiendas'):
        address='DireccionFiscal' if e['key']=='clientes' else 'Direccion'
        extra=f"CONCAT_WS(N' / ',NULLIF(m.{address},N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=m.CodigoUbigeo))"
    status='m.Estado' if e['key']=='versiones-tarifario' else "CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END"
    client_list=e['key'] in ('tiendas','contactos')
    client_name="CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END"
    search=[code,desc]
    if client_list:search+=['cliente.Codigo',client_name]
    if e['key']=='clientes':search+=['m.CodigoSAP']
    if e['key']=='unidades':search+=['CONVERT(NVARCHAR(40),m.Id)']
    if e['key']=='materiales':search+=['m.CodigoMaterialSAP']
    if e['key']=='partidas-tarifario':search+=['m.CodigoMaterialSAP','material.CodigoInterno']
    where='(@Patron IS NULL OR '+' OR '.join(x+" LIKE @Patron ESCAPE N'~'" for x in search)+')'
    where+=f" AND (@SoloActivos = 0 OR {e['selectable']})"
    if e['key']=='tarifarios':where+=' AND m.EsTarifarioMaterial=0'
    if e['key']=='versiones-tarifario':where+=' AND tarifa.EsTarifarioMaterial=0'
    if e['parent']:where+=f" AND (@PadreId IS NULL OR m.{e['parent']} = @PadreId)"
    body=permission(e,'CONSULTAR')+'''        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
'''
    cols=f'''m.{pk} AS Id, {code} AS Codigo, {desc} AS Descripcion, {status} AS Estado,
            {extra} AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion'''
    if client_list:cols+=f', cliente.Codigo AS ClienteCodigo, {client_name} AS ClienteNombre'
    list_joins=joins(e)+(" INNER JOIN dbo.Cliente AS cliente ON cliente.ClienteId=m.ClienteId" if client_list else '')
    client_select=', p.ClienteCodigo, p.ClienteNombre' if client_list else ''
    body+=f'''        SELECT {cols}
        INTO #Pagina FROM dbo.{table} AS m{list_joins}
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE {where};
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion{client_select}
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
'''
    procedure(f'spu_Mant{table}_Listar',f'Lista {table}, búsqueda SQL y página fija de veinte.',
        '    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,\n    @PadreId BIGINT = NULL, @SoloActivos BIT = 0',body)

    columns=[f'm.{pk} AS Id']+[f"m.{f['name']}" for f in e['fields']]+['m.VersionFila','m.FechaRegistro','creador.Nombres AS UsuarioRegistro','m.FechaActualizacion','editor.Nombres AS UsuarioActualizacion']
    if e['key']=='materiales':columns.append('ref3.Codigo AS UnidadCodigo')
    refs=''
    for i,field in enumerate(e['fields']):
        if field['lookup']:
            target=BY_KEY[field['lookup']]; alias='ref'+str(i)
            refs+=f" LEFT JOIN dbo.{target['table']} AS {alias} ON {alias}.{target['pk']} = m.{field['name']}"
            columns.append(label(target,alias)+f" AS {field['name']}Label")
    if e['zones']:
        ztable,zkey=e['zones']
        columns.append(f"JSON_QUERY((SELECT zona.ZonaId AS Id, catalogo.Codigo, catalogo.Descripcion FROM dbo.{ztable} AS zona INNER JOIN dbo.Zona AS catalogo ON catalogo.Id=zona.ZonaId WHERE zona.{zkey} = m.{pk} ORDER BY zona.ZonaId FOR JSON PATH)) AS Zonas")
    if e['key']=='partidas-tarifario':
        columns.append('padre.Estado AS VersionEstado')
        columns.append('(SELECT cabecera.EsTarifarioMaterial FROM dbo.Tarifario AS cabecera WHERE cabecera.TarifarioId=padre.TarifarioId) AS EsTarifarioMaterial')
    if e['key']=='tarifarios':columns.append("CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS v WHERE v.TarifarioId=m.TarifarioId AND v.Estado<>N'BORRADOR') THEN 1 ELSE 0 END AS BIT) AS TieneVersionesPublicadas")
    if e['key']=='clientes':columns.append(LINKS_SNAPSHOT[2:].replace('anterior.ClienteId','m.ClienteId'))
    selected=',\n            '.join(columns)
    body=permission(e,'CONSULTAR')+f'''        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT {selected}
        FROM dbo.{table} AS m{joins(e)}{refs}
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.{pk}=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'{table}' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'{table}' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
'''
    procedure(f'spu_Mant{table}_Obtener',f'Consulta {table}, referencias e historial paginado.',
        '    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1',body)

    fields=[f for f in e['fields'] if f['kind']!='ReadOnly']
    body=permission(e,'EDITAR')+'''        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
'''
    for field in fields:
        name,typ=field['name'],field['sql']
        if field['kind'] in ('AutoCode','Hidden'):
            body+=f'        DECLARE @{name} {typ};\n';continue
        raw=f"JSON_VALUE(@Datos,N'$.{name}')"
        length=re.search(r'NVARCHAR\((\d+)\)',typ)
        if length:
            body+=f"        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'{name}' AND LEN(dato.value)>{length.group(1)}) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: {name}.',1;\n"
            expr=f"NULLIF(LTRIM(RTRIM({raw})),N'')"
        elif typ=='BIT':expr=f"TRY_CONVERT(BIT,CASE {raw} WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE {raw} END)"
        elif typ=='DATE':expr=f"TRY_CONVERT(DATE,{raw},23)"
        else:expr=f'TRY_CONVERT({typ},{raw})'
        body+=f'        DECLARE @{name} {typ} = {expr};\n'
        if field['required']:body+=f"        IF @{name} IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: {name}.',1;\n"
        if not length:body+=f"        IF {raw} IS NOT NULL AND @{name} IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: {name}.',1;\n"
        if typ.startswith('DECIMAL'):body+=f"        IF @{name}<0 THROW 51207,N'[HYT_VALIDACION] El precio no puede ser negativo.',1;\n"
    body+='''        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
'''
    if e['key']=='partidas-tarifario':
        body+="        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.VersionId=@VersionId AND padre.Estado=N'BORRADOR') THROW 51208,N'[HYT_VALIDACION] Solo se pueden editar partidas de una versión borrador.',1;\n"
    if e['key']=='versiones-tarifario':
        body+="        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.TarifarioId=@TarifarioId) THROW 51208,N'[HYT_VALIDACION] El tarifario no existe.',1;\n"
    oldcols=', '.join(['anterior.'+f['name'] for f in e['fields']])
    if e['key']=='clientes':oldcols+=LINKS_SNAPSHOT
    if e['zones']:
        ztable,zkey=e['zones'];oldcols+=f', JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.{ztable} AS z WHERE z.{zkey}=anterior.{pk} FOR JSON PATH)) AS Zonas'
    body+=f'''        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT {oldcols} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.{table} AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.{pk}=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
'''
    if e['parent']:
        name=e['parent'];body+=f"        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.{name}'))<>@{name} THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;\n"
    if e['key']=='versiones-tarifario':body+="        IF @Id IS NOT NULL AND JSON_VALUE(@Anterior,N'$.Estado')<>N'BORRADOR' THROW 51211,N'[HYT_VALIDACION] La versión publicada es inmutable. Crea una nueva versión.',1;\n"
    for field in fields:
        if not field['lookup']:continue
        target=BY_KEY[field['lookup']]; name=field['name']; condition=target['selectable'].replace('m.','destino.')
        body+=f"        IF @{name} IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.{name}')),-1)<>@{name}) AND NOT EXISTS (SELECT 1 FROM dbo.{target['table']} AS destino WITH (HOLDLOCK) WHERE destino.{target['pk']}=@{name} AND {condition}) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: {name}.',1;\n"
    if e['key']=='clientes':
        body+='''        IF @TipoEntidad NOT IN (N'JURIDICA',N'NATURAL') THROW 51213,N'[HYT_VALIDACION] Tipo de entidad no válido.',1;
        IF (@TipoEntidad=N'JURIDICA' AND @RazonSocial IS NULL) OR (@TipoEntidad=N'NATURAL' AND (@Nombres IS NULL OR @Apellidos IS NULL)) THROW 51214,N'[HYT_VALIDACION] Completa la identificación del cliente.',1;
        SET @TipoIdentificacion=JSON_VALUE(@Anterior,N'$.TipoIdentificacion');
        IF @TipoEntidad=N'JURIDICA' BEGIN SET @Nombres=NULL; SET @Apellidos=NULL; END;
        ELSE SET @RazonSocial=NULL;
'''
    if e['key'] in ('clientes','tiendas'):
        body+="        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;\n"
    if e['key']=='clientes':body+=LINKS_VALIDATE
    if e['key']=='contactos':body+="        IF @TiendaId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId AND tienda.ClienteId=@ClienteId) THROW 51217,N'[HYT_VALIDACION] La tienda no pertenece al cliente.',1;\n"
    if e['key']=='tarifarios':body+="        IF @Id IS NOT NULL AND COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>COALESCE(@ClienteId,-1) AND EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS v WHERE v.TarifarioId=@Id AND v.Estado<>N'BORRADOR') THROW 51218,N'[HYT_VALIDACION] No se puede reasignar el cliente de un tarifario publicado.',1;\n"
    if any(f['name']=='VigenciaHasta' for f in fields):body+="        IF @VigenciaDesde IS NOT NULL AND @VigenciaHasta<@VigenciaDesde THROW 51219,N'[HYT_VALIDACION] Revisa las fechas de vigencia.',1;\n"
    if e['zones']:
        body+='''        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
'''
    if e.get('auto_prefix'):
        codefield=next(f['name'] for f in fields if f['kind']=='AutoCode')
        body+=f"        IF @Id IS NOT NULL SET @{codefield}=JSON_VALUE(@Anterior,N'$.{codefield}');\n"
        body+=f'''        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:{table}',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @{codefield}=N'{e['auto_prefix']}-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.{table} AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.{codefield}=@{codefield}) BREAK;
            END;
        END;
'''
    insertnames=[f['name'] for f in fields];insertvalues=['@'+n for n in insertnames]
    if e['key']=='versiones-tarifario':
        body+='''        DECLARE @NumeroVersion INT;
        SELECT @NumeroVersion=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
'''
        insertnames+=['NumeroVersion'];insertvalues+=['@NumeroVersion']
    body+=f'''        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.{table} ({', '.join(insertnames)}, UsuarioRegistro)
            VALUES ({', '.join(insertvalues)}, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.{table} SET {', '.join(n+'=@'+n for n in [f['name'] for f in fields if f['name']!=e['parent']])},
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE {pk}=@Id;
        END;
'''
    if e['zones']:
        ztable,zkey=e['zones'];body+=f'''        DELETE FROM dbo.{ztable} WHERE {zkey}=@Id;
        INSERT INTO dbo.{ztable} ({zkey},ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
'''
    if e['key']=='clientes':body+=LINKS_SAVE
    if e['key']=='partidas-tarifario':body+='        UPDATE dbo.TarifarioVersion SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@VersionId;\n'
    newcols=oldcols.replace('anterior.','posterior.')
    body+=f'''        SELECT @Posterior=(SELECT {newcols} FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.{table} AS posterior WHERE posterior.{pk}=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'{table}',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
'''
    procedure(f'spu_Mant{table}_Guardar',f'Guarda {table} con permisos, concurrencia, referencias y auditoría.',
        '    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),\n    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL',body,True)

save(BASE/'maintenance-generated.json',json.dumps(scripts,indent=2)+'\n')
print(f'{len(ENTITIES)} definiciones y {len(scripts)} procedimientos estáticos preparados, sin ejecutar SQL.')
