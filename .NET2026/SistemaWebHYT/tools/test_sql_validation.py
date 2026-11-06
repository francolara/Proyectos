"""Comprueba que el validador rechaza regresiones en copias temporales, sin SQL."""
from pathlib import Path
import shutil
import subprocess
import tempfile
import sys

project = Path(__file__).resolve().parents[1]
workspace = project.parents[1]
source = workspace / 'Basededatos' / 'Dbsishyt'
validator = Path(sys.argv[1]).resolve() if len(sys.argv)>1 else project / 'tools/Hyt.SqlValidation/bin/Debug/net10.0/Hyt.SqlValidation.dll'
if not validator.exists():
    raise RuntimeError('Compilar primero Hyt.SqlValidation en Debug.')

def validate(root):
    return subprocess.run(['dotnet', str(validator), str(root)], capture_output=True, text=True)

with tempfile.TemporaryDirectory(prefix='hyt-sql-validation-') as directory:
    target = Path(directory).resolve() / 'Dbsishyt'
    # Comprobación previa al borrado automático del directorio temporal.
    if not target.is_relative_to(Path(tempfile.gettempdir()).resolve()):
        raise RuntimeError('Destino temporal fuera del directorio previsto.')
    shutil.copytree(source, target)
    baseline = validate(target)
    if baseline.returncode != 0:
        raise RuntimeError(baseline.stdout + baseline.stderr)
    tests = [
        ('firma duplicada', 'StoreProcedure/spu_MaterialTarifas_Listar.sql', '-- Create date:   05/11/2026', '-- Create date:   05/11/2026\r\n-- Create date:   05/11/2026', 'Firma incorrecta'),
        ('fecha imposible', 'StoreProcedure/spu_MaterialTarifas_Listar.sql', '-- Create date:   05/11/2026', '-- Create date:   32/11/2026', 'Firma incorrecta'),
        ('sintaxis inválida', 'Tablas/Cliente.sql', 'CREATE TABLE', 'CREATE TABLA', 'Cliente.sql:'),
        ('referencia inexistente', 'Tablas/Tienda.sql', 'REFERENCES dbo.Cliente(ClienteId)', 'REFERENCES dbo.NoExiste(ClienteId)', 'Dependencia fuera de orden'),
        ('segunda guía permitida', 'Tablas/GuiaConformidad.sql', 'CONSTRAINT UQ_Guia_Presupuesto UNIQUE', '', 'Invariante ausente'),
        ('SELECT * prohibido', 'StoreProcedure/spu_ListarClientes.sql', 'SELECT COUNT_BIG(1)', 'SELECT * --\n        SELECT COUNT_BIG(1)', 'SELECT *'),
        ('columna inexistente en SP', 'StoreProcedure/spu_MantCliente_Listar.sql', 'm.Codigo AS Codigo', 'm.ColumnaInexistente AS Codigo', 'Columna referida ausente'),
    ]
    for label, relative, old, new, expected in tests:
        path = target / relative
        original = path.read_bytes()
        text = original.decode('utf-8')
        if old not in text:
            raise RuntimeError('Caso de prueba no aplicable: ' + label)
        path.write_bytes(text.replace(old, new, 1).encode('utf-8'))
        result = validate(target)
        path.write_bytes(original)
        if result.returncode == 0 or expected not in result.stderr:
            raise RuntimeError(f'No se detectó {label}: {result.stdout}\n{result.stderr}')
        print('Detectado correctamente: ' + label)
print('Validación base y 7 regresiones comprobadas offline; no se ejecutó SQL.')
