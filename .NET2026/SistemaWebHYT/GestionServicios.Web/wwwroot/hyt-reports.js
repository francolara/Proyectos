// Firma: FRANCO LARA - 03/11/2026 | Impresión de reporte HTML.
// Firma: FRANCO LARA - 05/11/2026 | Actualiza filtros de reporte.
// Firma: FRANCO LARA - 06/11/2026 | Cambiar fuente limpia estado dependiente y vuelve a consultar; conserva búsqueda y fechas.
const reportKind=document.querySelector('#report-kind');
reportKind?.addEventListener('change',()=>{const form=reportKind.form;form.querySelector('[name=state]').value='';for(const name of ['pendingOrder','pendingConfirmation','pendingInvoice']){const input=form.querySelector(`[name=${name}]`);if(input)input.checked=false;}form.requestSubmit();});
