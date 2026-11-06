"""Compara fuentes SP con SQL local; solo lectura, sin datos ni secretos en salida."""
from pathlib import Path
import re
import subprocess

root = Path(__file__).resolve().parents[3] / 'Basededatos/Dbsishyt/StoreProcedure'
count = 0
for source in sorted(root.glob('*.sql')):
    name = source.stem
    if not re.fullmatch(r'spu_[A-Za-z0-9_]+', name):
        raise RuntimeError('Nombre de módulo inesperado.')
    text = source.read_text(encoding='utf-8-sig')
    expected = text[text.index('CREATE OR ALTER PROCEDURE'):].strip()
    expected = re.split(r'^\s*GO\s*$', expected, maxsplit=1, flags=re.I | re.M)[0].strip()
    query = "SET NOCOUNT ON; SELECT definition FROM sys.sql_modules WHERE object_id=OBJECT_ID(N'dbo." + name + "');"
    result = subprocess.run(['sqlcmd', '-S', r'localhost\SQLEXPRESS', '-E', '-C', '-I', '-b', '-f', '65001', '-d', 'SistemaWebHYT', '-y', '0', '-w', '65535', '-Q', query], capture_output=True)
    if result.returncode:
        raise RuntimeError('No se pudo consultar el módulo ' + name)
    actual = result.stdout.decode('utf-8-sig').strip()
    normalize = lambda value: re.sub(r'\s+', ' ', re.sub(r'^CREATE(?: OR ALTER)? PROCEDURE', 'CREATE PROCEDURE', value, flags=re.I)).strip()
    if normalize(actual) != normalize(expected):
        raise RuntimeError('Definición instalada distinta: ' + name)
    count += 1
print(f'OK SQL local: {count} procedimientos instalados coinciden con sus fuentes.')
