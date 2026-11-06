"""Valida los cuatro reportes HTML impresos desde el host aislado de QA."""
from pathlib import Path
import re
import pdfplumber
from pypdf import PdfReader

root = Path(__file__).resolve().parents[1] / 'entregables/preview'
for kind in ['ganados', 'oc_completadas', 'oc_con_acta', 'oc_pendientes_factura']:
    path = root / f'reporte-{kind}.pdf'
    reader = PdfReader(path)
    texts = [page.extract_text() or '' for page in reader.pages]
    full = '\n'.join(texts)
    prefix = 'PPTO' if kind == 'ganados' else 'OC'
    numbers = re.findall(rf'{prefix}-(\d{{4}})A', full)
    assert sorted(map(int, numbers)) == list(range(1, 21)), (kind, 'filas perdidas o duplicadas')
    assert '9,676.00' in full and '41 documentos' in full, (kind, 'totales globales')
    assert 'Página del sistema 1' in full, (kind, 'alcance de impresión')
    for index, page in enumerate(reader.pages):
        assert abs(float(page.mediabox.width) - 842) < 2 and abs(float(page.mediabox.height) - 595) < 2, (kind, 'A4 horizontal')
        assert 'Fecha de registro' in texts[index] and 'Página del sistema 1' in texts[index], (kind, 'encabezado repetido')
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            assert all(char['x0'] >= -1 and char['x1'] <= page.width + 1 and char['top'] >= -1 and char['bottom'] <= page.height + 1 for char in page.chars), (kind, 'texto fuera de página')
    print(f'OK PDF: {kind}, {len(reader.pages)} páginas A4 horizontal, 20 filas completas y encabezados repetidos.')
