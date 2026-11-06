"""Empaqueta etapas 5/6/7/8 desde sus manifiestos y verifica huellas; no ejecuta SQL."""
from pathlib import Path
import hashlib
import json
import re
ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parents[1]/'Basededatos/Dbsishyt'
OUTPUT=ROOT/'entregables'
OUTPUT.mkdir(exist_ok=True)
def save(path,text):path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
for stage in (5,6,7,8):
    manifest=json.loads((BASE/f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))
    scripts=manifest['scripts']
    if len(scripts)!=len(set(scripts)):raise RuntimeError('Fuentes duplicadas.')
    parts=[f'''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Paquete incremental HYT etapa {stage}; aplicar después de etapa {stage-1}.
-- =============================================
-- Las fuentes individuales están aplicadas en la base local de desarrollo.
''']
    hashes=[]
    updated_today=False
    updated_nov6=False
    for relative in scripts:
        source=(BASE/relative).resolve()
        if not source.is_relative_to(BASE.resolve()):raise RuntimeError('Fuente fuera del directorio SQL.')
        raw=source.read_bytes()
        updated_today=updated_today or b'05/11/2026' in raw
        updated_nov6=updated_nov6 or b'06/11/2026' in raw
        # El paquete conserva una firma por fecha; las firmas individuales
        # permanecen en las fuentes y sus huellas identifican cada archivo.
        text=re.sub(r'\A(?:[ \t]*(?:--[^\r\n]*)?\r?\n)+','',raw.decode('utf-8'))
        parts.append(f'\n-- Fuente: {relative}\n{text}\n')
        hashes.append({'path':relative,'sha256':hashlib.sha256(raw).hexdigest()})
    target=OUTPUT/f'ETAPA_{stage}_HYT.sql'
    if updated_today:
        parts.insert(1,'-- =============================================\n-- Author:        FRANCO LARA\n-- Create date:   05/11/2026\n-- Description:   Actualiza consultas y filtros del dashboard, navegación comercial y resultados vacíos según fuentes de esta etapa.\n-- =============================================\n')
    if updated_nov6:parts.insert(1,'-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración mensual de presupuesto en consultas, reportes e impresiones; preserva claves internas.\n')
    save(target,''.join(parts))
    hashes.append({'path':target.name,'sha256':hashlib.sha256(target.read_bytes()).hexdigest()})
    save(OUTPUT/f'SHA256_ETAPA_{stage}.json',json.dumps(hashes,indent=2)+'\n')
    for item in hashes:
        candidate=target if item['path']==target.name else BASE/item['path']
        if hashlib.sha256(candidate.read_bytes()).hexdigest()!=item['sha256']:raise RuntimeError('Huella no coincide.')
    print(f'Etapa {stage}: {len(scripts)} fuentes; paquete y huellas verificados.')
