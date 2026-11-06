"""Definiciones revisables de mantenimientos. Solo genera archivos; nunca conecta SQL."""
def f(name, label, sql='NVARCHAR(150)', required=False, kind='Text', lookup=None, default=None, catalog=None):
    return dict(name=name, label=label, sql=sql, required=required, kind=kind, lookup=lookup, default=default, catalog=catalog)

def ref(name, label, lookup, required=True, sql='BIGINT'):
    return f(name,label,sql,required,'Reference',lookup)

ACTIVE = f('Activo','Activo','BIT',True,'Boolean',default='true')
ENTITIES = []
def entity(key, title, table, pk, permission, fields, code, description, parent=None, zones=None, selectable='m.Activo = 1', extra="N''"):
    ENTITIES.append(dict(key=key,title=title,table=table,pk=pk,permission=permission,fields=fields,code=code,description=description,parent=parent,zones=zones,selectable=selectable,extra=extra))

entity('clientes','Clientes','Cliente','ClienteId','CLIENTES',[
    f('Codigo','Código','NVARCHAR(40)',True),f('CodigoSAP','Código SAP','NVARCHAR(60)'),f('TipoEntidad','Tipo de entidad','NVARCHAR(20)',True,'Select',default='JURIDICA',catalog='EntityType'),
    f('TipoIdentificacion','Tipo de documento','NVARCHAR(20)',kind='Hidden'),f('NumeroIdentificacion','Número de documento','NVARCHAR(30)',True),
    f('RazonSocial','Razón social','NVARCHAR(200)'),f('Nombres','Nombres'),f('Apellidos','Apellidos'),
    f('DireccionFiscal','Dirección fiscal','NVARCHAR(300)'),f('CodigoUbigeo','Ubigeo','NVARCHAR(6)',kind='Ubigeo'),f('UrlCargaDocumentos','Enlace anterior','NVARCHAR(2048)',kind='Hidden'),ACTIVE
], 'm.Codigo', "CASE WHEN m.TipoEntidad=N'NATURAL' THEN CONCAT(m.Nombres, N' ', m.Apellidos) ELSE m.RazonSocial END")
entity('tiendas','Tiendas','Tienda','TiendaId','TIENDAS',[
    ref('ClienteId','Cliente','clientes'),f('Codigo','Código','NVARCHAR(40)',True),f('Nombre','Nombre de tienda','NVARCHAR(200)',True),
    f('Direccion','Dirección','NVARCHAR(300)'),f('CodigoUbigeo','Ubigeo','NVARCHAR(6)',kind='Ubigeo'),f('CentroSAPFICO','Centro SAP FICO','NVARCHAR(50)'),ref('ZonaId','Zona','zonas',False,'INT'),ACTIVE
], 'm.Codigo','m.Nombre',parent='ClienteId')
entity('contactos','Contactos','ContactoCliente','ContactoId','CLIENTES',[
    ref('ClienteId','Cliente','clientes'),ref('TiendaId','Tienda (opcional)','tiendas',False),f('Nombre','Nombre','NVARCHAR(150)',True),
    f('Telefono','Teléfono','NVARCHAR(50)'),f('Cargo','Cargo','NVARCHAR(100)'),f('Correo','Correo electrónico','NVARCHAR(256)',kind='Email'),ACTIVE
], 'CONVERT(NVARCHAR(40), m.ContactoId)','m.Nombre',parent='ClienteId')
entity('proveedor-cliente','Datos de HYT por cliente','ProveedorCliente','Id','CLIENTES',[
    ref('ClienteId','Cliente','clientes'),f('CodigoSAPProveedor','Código SAP del proveedor','NVARCHAR(60)'),f('ContratoMarco','Contrato marco','NVARCHAR(60)'),
    f('VigenciaDesde','Vigencia desde','DATE',kind='Date'),f('VigenciaHasta','Vigencia hasta','DATE',kind='Date'),ACTIVE
], "CONVERT(NVARCHAR(40), m.Id)", "COALESCE(m.ContratoMarco, N'Datos del proveedor')",parent='ClienteId')
for key,title,table,permission in [('unidades','Unidades de medida','UnidadMedida','UNIDADES'),('zonas','Zonas','Zona','ZONAS'),('especialidades','Especialidades','Especialidad','MATERIALES'),('tipos-servicio','Tipos de servicio','TipoServicio','MATERIALES')]:
    entity(key,title,table,'Id',permission,[f('Codigo','Código','NVARCHAR(30)',True),f('Descripcion','Descripción','NVARCHAR(150)',True),ACTIVE], 'm.Codigo','m.Descripcion')
entity('materiales','Materiales y servicios','ServicioMaterial','MaterialId','MATERIALES',[
    f('CodigoInterno','Código interno','NVARCHAR(60)',True),f('CodigoMaterialSAP','Código de material SAP','NVARCHAR(60)'),
    f('Descripcion','Descripción','NVARCHAR(1500)',True,'LongText'),ref('UnidadMedidaId','Unidad de medida','unidades',True,'INT'),ref('EspecialidadId','Especialidad','especialidades',False,'INT'),ACTIVE
], 'm.CodigoInterno','m.Descripcion',zones=('MaterialZona','MaterialId'),extra="(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), z.Descripcion), N', ') FROM dbo.MaterialZona AS mz INNER JOIN dbo.Zona AS z ON z.Id = mz.ZonaId WHERE mz.MaterialId = m.MaterialId)")
entity('tarifarios','Tarifarios','Tarifario','TarifarioId','TARIFARIOS',[
    f('Codigo','Código','NVARCHAR(40)',True),f('Nombre','Nombre','NVARCHAR(200)',True),ref('ClienteId','Cliente (vacío: genérico)','clientes',False),ACTIVE
], 'm.Codigo','m.Nombre')
entity('versiones-tarifario','Versiones de tarifario','TarifarioVersion','VersionId','TARIFARIOS',[
    ref('TarifarioId','Tarifario','tarifarios'),f('NumeroVersion','Número de versión','INT',kind='ReadOnly'),
    f('VigenciaDesde','Vigencia desde','DATE',kind='Date'),f('VigenciaHasta','Vigencia hasta','DATE',kind='Date'),
    f('CodigoMonedaPredeterminada','Moneda predeterminada','NVARCHAR(3)',True,'Select',default='PEN',catalog='Currency'),
    f('CodigoTipoIGVPredeterminado','Tratamiento IGV predeterminado','NVARCHAR(10)',True,'Select',catalog='TaxTreatment'),
    f('Estado','Estado','NVARCHAR(20)',kind='ReadOnly',default='BORRADOR')
], 'CONVERT(NVARCHAR(40), m.NumeroVersion)', "CONCAT(tarifa.Nombre, N' · Versión ', m.NumeroVersion)",parent='TarifarioId',zones=('TarifarioVersionZona','VersionId'),selectable="m.Estado = N'BORRADOR'",extra='m.Estado')
entity('partidas-tarifario','Partidas del tarifario','TarifarioPartida','PartidaId','TARIFARIOS',[
    ref('VersionId','Versión de tarifario','versiones-tarifario'),ref('MaterialId','Material / servicio','materiales'),
    f('Posicion','Posición','NVARCHAR(60)'),f('CodigoMaterialSAP','Código de material SAP','NVARCHAR(60)'),
    ref('TipoServicioId','Tipo de servicio','tipos-servicio',False,'INT'),f('NumeroNecesidad','Número de necesidad','NVARCHAR(60)'),
    f('CodigoMoneda','Moneda de referencia','NVARCHAR(3)',True,'Select',default='PEN',catalog='Currency'),
    f('CodigoTipoIGV','Tratamiento IGV de los precios','NVARCHAR(10)',True,'Select',catalog='TaxTreatment'),
    f('PrecioPEN','Precio en soles (PEN)','DECIMAL(19,6)',kind='Decimal'),f('PrecioUSD','Precio en dólares (USD), opcional','DECIMAL(19,6)',kind='Decimal'),ACTIVE
], "COALESCE(m.Posicion, CONVERT(NVARCHAR(40), m.PartidaId))", 'material.Descripcion',parent='VersionId',zones=('TarifarioPartidaZona','PartidaId'),extra="(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), z.Descripcion), N', ') FROM dbo.TarifarioPartidaZona AS pz INNER JOIN dbo.Zona AS z ON z.Id = pz.ZonaId WHERE pz.PartidaId = m.PartidaId)")

# Los códigos internos son exclusivamente del servidor. Unidad conserva Codigo como abreviatura compatible con históricos.
AUTO_PREFIXES={'clientes':'CLI','tiendas':'TIE','zonas':'ZON','especialidades':'ESP','tipos-servicio':'TS','materiales':'MAT','tarifarios':'TAR'}
for e in ENTITIES:
    e['auto_prefix']=AUTO_PREFIXES.get(e['key'])
    if e['auto_prefix']:
        for field in e['fields']:
            if field['name'] in ('Codigo','CodigoInterno'):
                field['kind']='AutoCode';field['required']=False;field['label']='Código interno'
    if e['key']=='unidades':
        for field in e['fields']:
            if field['name']=='Codigo':field['label']='Abreviatura'

BY_KEY = {e['key']:e for e in ENTITIES}
