"""Verifica impresión de HTML en Chrome. Solo lee los PDF de prueba."""
from pathlib import Path
import json
import re
import pdfplumber
from pypdf import PdfReader

root=Path(__file__).resolve().parents[1]/'entregables/preview'
cases=['nuevo-una','anterior-una','nuevo-varias','anterior-varias','nuevo-vacio','nuevo-usd','nuevo-descripcion-larga']
result=[]
for case in cases:
    path=root/(case+'.pdf')
    reader=PdfReader(path)
    texts=[page.extract_text() or '' for page in reader.pages]
    full='\n'.join(texts)
    assert all(abs(float(page.mediabox.width)-595)<2 and abs(float(page.mediabox.height)-842)<2 for page in reader.pages),case+' no es A4'
    if case in ['nuevo-una','anterior-una','nuevo-vacio','nuevo-usd']:assert len(reader.pages)==1,case+' requiere una página para este caso de prueba'
    if 'varias' in case:
        assert len(reader.pages)>1
        numbers=[int(number) for number in re.findall(r'Servicio\s+(\d+)\.',full)]
        assert sorted(numbers)==list(range(1,41)),case+' perdió o duplicó partidas'
        assert all('Descripción' in page for page in texts[:-1]),case+' perdió encabezados de tabla'
        assert '4,720.00' in full,case+' total incorrecto'
    if case=='nuevo-descripcion-larga':assert 'FIN-DESCRIPCION' in full,case+' descripción incompleta'
    if case=='nuevo-vacio':assert 'todavía no contiene partidas' in full and 'Pendiente' in full
    if case=='nuevo-usd':assert 'USD' in full and '$ 118.00' in full and 'incluyen IGV' in full
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            assert not any(char['x0'] < -1 or char['x1'] > page.width+1 or char['top'] < -1 or char['bottom'] > page.height+1 for char in page.chars),case+' texto fuera de página'
    result.append({'case':case,'pages':len(reader.pages),'a4':True,'text_complete':True})
    print(f'OK PDF: {case}, {len(reader.pages)} página(s), texto completo y dentro del papel.')
(root/'report-pdf-validation.json').write_bytes((json.dumps(result,indent=2)+'\n').replace('\n','\r\n').encode('utf-8'))
