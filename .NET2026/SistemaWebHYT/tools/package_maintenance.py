"""Manifiestos y paquetes incrementales 3/4; sin ejecución SQL."""
import hashlib
import json
import re
from pathlib import Path
from maintenance_schema import ENTITIES

PROJECT = Path(__file__).resolve().parents[1]
BASE = Path(__file__).resolve().parents[3] / 'Basededatos/Dbsishyt'
OUTPUT = PROJECT / 'entregables'
def save(path, text):
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))

groups = {3: ['Tablas/UbigeoDepartamentos.sql','Tablas/UbigeoProvincias.sql','Tablas/UbigeoDistritos.sql','Script/024_UbigeosCatalogo.sql','Script/025_DireccionesYConsideraciones.sql','Script/028_UbigeoConfiguracionHYT.sql','Script/011_EstadosMantenimientos.sql', 'Script/020_CodigosInternosAutomaticos.sql', 'Script/021_CodigoSAPCliente.sql', 'Tablas/ClienteEnlaceDocumento.sql', 'Script/022_EnlacesDocumentosCliente.sql', 'StoreProcedure/spu_MantCatalogos.sql'], 4: ['Script/023_TarifarioMaterialDirecto.sql']}
for entity in ENTITIES:
    stage = 4 if entity['permission'] == 'TARIFARIOS' else 3
    groups[stage] += [f"StoreProcedure/spu_Mant{entity['table']}_{action}.sql" for action in ['Listar','Obtener','Guardar']]
groups[4] += ['StoreProcedure/spu_MaterialTarifas_Listar.sql','StoreProcedure/spu_MaterialTarifa_Guardar.sql']
groups[4] += ['StoreProcedure/spu_MantTarifarioVersion_Publicar.sql', 'StoreProcedure/spu_MantTarifarioVersion_Clonar.sql']
groups[4] += ['StoreProcedure/spu_MantTasaIGV_Listar.sql', 'StoreProcedure/spu_MantTasaIGV_Agregar.sql']
signature = '''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   {description}
-- =============================================
'''
for stage, scripts in groups.items():
    grants = f'Script/{stage+9:03}_PermisosEtapa{stage}.sql'
    sql = signature.format(description=f'Permisos limitados de mantenimientos, etapa {stage}.') + 'USE [SistemaWebHYT];\nGO\nSET NOCOUNT ON;\n'
    for source in scripts:
        if source.startswith('StoreProcedure/'):
            sql += f'GRANT EXECUTE ON OBJECT::dbo.{Path(source).stem} TO hyt_app_executor;\n'
    if stage==4:sql=sql.replace('USE [SistemaWebHYT];','-- Firma: FRANCO LARA | Create date: 05/11/2026 | Agrega consulta de precios por material con permisos existentes.\nUSE [SistemaWebHYT];',1)
    save(BASE / grants, sql + 'GO\n')
    scripts.append(grants)
    save(BASE / f'manifest-etapa{stage}.json', json.dumps({'database':'SistemaWebHYT','requires':f'etapa{stage-1}','scripts':scripts}, indent=2)+'\n')
    parts = [signature.format(description=f'Entrega incremental HYT etapa {stage}, para ejecución manual.') + f'-- Aplicar después de etapa {stage-1}. No ejecutado automáticamente.\n']
    if any(b'05/11/2026' in (BASE / source).read_bytes() for source in scripts):parts.append('-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos internos automáticos y precios gestionados desde materiales.\n')
    if any(b'06/11/2026' in (BASE / source).read_bytes() for source in scripts):parts.append('-- Firma: FRANCO LARA | Create date: 06/11/2026 | Ubigeo de configuración HYT y demás fuentes de esta etapa.\n')
    hashes = []
    for relative in scripts:
        raw = (BASE / relative).read_bytes()
        text = re.sub(r'\A(?:[ \t]*(?:--[^\r\n]*)?\r?\n)+', '', raw.decode('utf-8'))
        parts.append(f'\n-- Fuente: {relative}\n{text}\n')
        hashes.append({'path':relative,'sha256':hashlib.sha256(raw).hexdigest()})
    bundle = OUTPUT / f'ETAPA_{stage}_HYT.sql'
    save(bundle, ''.join(parts))
    hashes.append({'path':bundle.name,'sha256':hashlib.sha256(bundle.read_bytes()).hexdigest()})
    save(OUTPUT / f'SHA256_ETAPA_{stage}.json', json.dumps(hashes, indent=2)+'\n')
    print(f'Etapa {stage}: {len(scripts)} fuentes, paquete y huellas preparados.')

prior = []
for stage in [1,2]:
    prior += json.loads((BASE / f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))['scripts']
manifest = json.loads((BASE / 'manifest.json').read_text(encoding='utf-8'))
manifest['scripts'] = prior + groups[3] + groups[4] + [source for stage in range(5,10) if (BASE/f'manifest-etapa{stage}.json').exists() for source in json.loads((BASE/f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))['scripts']]
save(BASE / 'manifest.json', json.dumps(manifest, indent=2)+'\n')
