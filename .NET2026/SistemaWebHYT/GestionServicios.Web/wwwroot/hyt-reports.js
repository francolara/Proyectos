// Firma: FRANCO LARA - 03/11/2026 | Impresión de vista HTML del reporte.
document.querySelector('[data-report-print]')?.addEventListener('click',()=>window.print());

// Firma: FRANCO LARA - 05/11/2026 | El filtro de seguimiento solo se aplica a reportes de OC.
(()=>{const kind=document.querySelector("#report-kind"),situation=document.querySelector("#report-situation");if(!kind||!situation)return;const update=()=>{situation.disabled=kind.value==="GANADOS";if(situation.disabled)situation.value="";};kind.addEventListener("change",update);update();})();
