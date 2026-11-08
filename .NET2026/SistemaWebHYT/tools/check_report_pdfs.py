"""Verifica impresión de HTML en Chrome. Solo lee los PDF de prueba."""
from pathlib import Path
import json
import re
import pdfplumber
from pypdf import PdfReader

root=Path(__file__).resolve().parents[1]/'entregables/preview/anterior-contacto-tienda'
cases=['nuevo-una','anterior-una','nuevo-varias','anterior-varias','nuevo-vacio','nuevo-usd','nuevo-descripcion-larga','anterior-vacio','anterior-usd','anterior-descripcion-larga']
result=[]
for case in cases:
    path=root/(case+'.pdf')
    reader=PdfReader(path)
    texts=[page.extract_text() or '' for page in reader.pages]
    full='\n'.join(texts)
    assert all(abs(float(page.mediabox.width)-595)<2 and abs(float(page.mediabox.height)-842)<2 for page in reader.pages),case+' no es A4'
    if case in ['nuevo-una','anterior-una','nuevo-vacio','nuevo-usd','anterior-vacio','anterior-usd']:assert len(reader.pages)==1,case+' requiere una página para este caso de prueba'
    if 'varias' in case:
        assert len(reader.pages)>1
        numbers=[int(number) for number in re.findall(r'Servicio\s+(\d+)\.',full)]
        assert sorted(numbers)==list(range(1,41)),case+' perdió o duplicó partidas'
        assert all(('DESCRIPCION' in page or 'DESCRIPCIÓN' in page or 'Descripción' in page) for page in texts[:-1]),case+' perdió encabezados de tabla'
        assert ('4,000.00' if case.startswith('anterior') else '4,720.00') in full,case+' total incorrecto'
    if case.endswith('descripcion-larga'):assert 'FIN-DESCRIPCION' in full,case+' descripción incompleta'
    if case.endswith('vacio'):assert 'todavía no contiene partidas' in full and 'Pendiente' in full
    if case.endswith('usd'):assert 'USD' in full and ('$ 100.00' if case.startswith('anterior') else '$ 118.00') in full
    if case.startswith('anterior'):
        assert 'COSTO FINAL SIN IGV' in full and 'PPTO-' in full
        assert 'T001' not in full and 'IGV 18%' not in full
        assert 'Coordinador/Encargado:' in full and 'Contacto de tienda' in full
        assert not any(label in full for label in ['Revisión','Generado','Los precios unitarios'])
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            assert not any(char['x0'] < -1 or char['x1'] > page.width+1 or char['top'] < -1 or char['bottom'] > page.height+1 for char in page.chars),case+' texto fuera de página'
    result.append({'case':case,'pages':len(reader.pages),'a4':True,'text_complete':True})
    print(f'OK PDF: {case}, {len(reader.pages)} página(s), texto completo y dentro del papel.')
(root/'report-pdf-validation.json').write_bytes((json.dumps(result,indent=2)+'\n').replace('\n','\r\n').encode('utf-8'))
