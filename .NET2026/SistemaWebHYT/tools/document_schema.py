"""Lista cerrada para documentos de seguimiento. No ejecuta SQL."""
def field(name,label,sql='NVARCHAR(100)',required=False,kind='Text'):
    return dict(name=name,label=label,sql=sql,required=required,kind=kind)

DOCUMENTS=[
 dict(key='oc-cliente',code='OC',title='Orden de Compra - Cliente',table='OrdenCompraCliente',pk='OrdenCompraId',permission='OC_CLIENTE',parent='PresupuestoId',state='REGISTRADA',number='NumeroRegistroInterno',filetable='OrdenCompraArchivo',filekey='OrdenCompraId',fields=[
     field('RevisionId','Revisión del presupuesto','BIGINT',True,'Revision'),field('NumeroOCCliente','N° OC del cliente',required=True),
     field('FechaEmisionOC','Fecha de emisión OC','DATE',True,'Date'),field('FechaRecepcionOC','Fecha de recepción OC','DATE',True,'Date'),
     field('Importe','Importe de la OC','DECIMAL(19,2)',False,'Money'),field('Observacion','Observación','NVARCHAR(2000)',False,'LongText')]),
 dict(key='guias',code='GUIA',title='Guia (Acta de conformidad)',table='GuiaConformidad',pk='GuiaId',permission='GUIAS',parent='PresupuestoId',state='REGISTRADA',number='NumeroActa',filetable='GuiaArchivo',filekey='GuiaId',fields=[
     field('FechaActa','Fecha de la guía','DATE',True,'Date'),field('NumeroActa','N° guía / acta',required=True),field('Observacion','Observación','NVARCHAR(2000)',False,'LongText')]),
 dict(key='confirmaciones',code='CONFIRMACION',title='Confirmación del cliente',table='ConfirmacionFacturacion',pk='ConfirmacionId',permission='CONFIRMACIONES',parent='OrdenCompraId',state='VIGENTE',number='NumeroConfirmacion',filetable='ConfirmacionArchivo',filekey='ConfirmacionId',fields=[
     field('FechaConfirmacion','Fecha de confirmación','DATE',True,'Date'),field('NumeroConfirmacion','N° confirmación',required=True),field('RepresentanteCliente','Representante del cliente','NVARCHAR(150)'),field('Observacion','Observación','NVARCHAR(2000)',False,'LongText')]),
 dict(key='ejecucion',code='EJECUCION',title='Ejecución del trabajo',table='Ejecucion',pk='EjecucionId',permission='EJECUCION',parent='PresupuestoId',state='PENDIENTE',number='EjecucionId',filetable=None,filekey=None,fields=[
     field('Responsable','Responsable del trabajo','NVARCHAR(150)'),field('FechaProgramada','Fecha programada','DATE',False,'Date'),field('FechaInicioReal','Fecha de inicio real','DATE',False,'Date'),field('FechaTerminoReal','Fecha de término real','DATE',False,'Date'),field('Descripcion','Descripción del trabajo realizado','NVARCHAR(2000)',False,'LongText')])
]
