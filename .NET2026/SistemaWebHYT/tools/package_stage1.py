"""Genera entrega SQL para revisión/ejecución manual. Nunca abre conexiones."""
from pathlib import Path
import hashlib
import json
import re

ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / 'Basededatos' / 'Dbsishyt'
OUTPUT = Path(__file__).resolve().parents[1] / 'entregables'
manifest = json.loads((BASE / 'manifest-etapa1.json').read_text(encoding='utf-8'))
parts = ['''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Entrega inicial HYT, esquema y catalogos para ejecucion manual.
-- =============================================
-- Revisar LEEME_DESPLIEGUE.md. Este archivo no se ejecuta automaticamente.
-- Preparado para una base nueva SistemaWebHYT, no para migrar otros esquemas.
''']
if any(b'05/11/2026' in (BASE / item).read_bytes() for item in manifest['scripts']):
    parts.append('-- Firma: FRANCO LARA | Create date: 05/11/2026 | Código SAP opcional en el maestro de clientes.\n')
if any(b'06/11/2026' in (BASE / item).read_bytes() for item in manifest['scripts']):
    parts.append('-- Firma: FRANCO LARA | Create date: 06/11/2026 | Nombres completos del tipo de servicio en presupuesto y ubigeo de HYT.\n')
hashes = []
for relative in manifest['scripts']:
    path = BASE / relative
    raw = path.read_bytes()
    text = raw.decode('utf-8')
    # La firma del paquete es única; las firmas de cada fuente se conservan allí.
    text = re.sub(r'\A(?:[ \t]*(?:--[^\r\n]*)?\r?\n)+', '', text)
    parts.append(f'\n-- Fuente: {relative}\n{text}\n')
    hashes.append({'path': relative, 'sha256': hashlib.sha256(raw).hexdigest()})
OUTPUT.mkdir(exist_ok=True)
bundle = OUTPUT / 'ETAPA_1_HYT.sql'
bundle.write_bytes(''.join(parts).replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8'))
hashes.append({'path': bundle.name, 'sha256': hashlib.sha256(bundle.read_bytes()).hexdigest()})
(OUTPUT / 'SHA256.json').write_bytes((json.dumps(hashes, indent=2) + '\n').replace('\n', '\r\n').encode('utf-8'))
print(f'Paquete preparado: {len(manifest["scripts"])} scripts. No se ha ejecutado SQL.')
