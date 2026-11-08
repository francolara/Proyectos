// Firma: FRANCO LARA - 06/11/2026 | Motivo visible y obligatorio solo al observar, rechazar o cancelar presupuesto; consultas de adjuntos y comprobante con doble clic o teclado.
// Firma: FRANCO LARA - 03/11/2026 | Captura documental en popup compartido, foco, validación y bloqueo de doble envío.
// Firma: FRANCO LARA - 05/11/2026 | Mantiene la espera común en el retorno tras guardar documentos.
import {setDialog} from './hyt-ui.js';
const dialog=document.querySelector('#document-dialog');
const body=dialog?.querySelector('[data-document-body]');
let loading=false;
function updateBudgetReason(form){
    if(!form)return;
    const state=form.querySelector('[data-budget-state]')?.value;
    const needed=['OBSERVADO','RECHAZADO','CANCELADO'].includes(state);
    const field=form.querySelector('[data-budget-state-reason]');if(!field)return;
    field.hidden=!needed;const input=field.querySelector('textarea');input.disabled=!needed;input.required=needed;
}
document.querySelectorAll('[data-budget-state-form]').forEach(updateBudgetReason);
document.addEventListener('change',event=>{if(event.target.matches('[data-budget-state]'))updateBudgetReason(event.target.closest('form'));});
async function openDocument(event){
    const link=event.target.closest('[data-document-open],[data-document-doubleclick]');if(!link||!dialog)return;
    const doubleClick=link.dataset.documentDoubleclick?.toLowerCase()==='true';
    if(event.type==='dblclick'&&!doubleClick)return;
    if(doubleClick&&event.type==='click'&&event.detail!==0){event.preventDefault();return;}
    event.preventDefault();if(loading)return;loading=true;
    dialog.querySelector("h2").textContent="Registro documental";window.hytDisposeDates?.(body);const loadingMessage=document.createElement('p');loadingMessage.textContent='Cargando registro…';body.replaceChildren(loadingMessage);setDialog(dialog,true,true);
    try{
        const url=new URL(link.href,location.href);if(url.origin!==location.origin)throw new Error('Enlace no válido.');url.searchParams.set('popup','true');
        const response=await fetch(url,{headers:{'X-Requested-With':'XMLHttpRequest'}});
        if(!response.ok||response.redirected)throw new Error(response.redirected?'La sesión o los permisos cambiaron. Vuelve a abrir el listado.':await response.text());
        body.innerHTML=await response.text();dialog.querySelector("h2").textContent=body.querySelector("[data-document-title]")?.dataset.documentTitle??"Registro documental";window.hytInitDates?.(body);body.querySelectorAll("[data-budget-state-form]").forEach(updateBudgetReason);setDialog(dialog,true,false);body.querySelector('input:not([type=hidden]),select,textarea')?.focus();
    }catch(error){const message=document.createElement('p');message.setAttribute('role','alert');message.textContent=error.message.includes('<')?'No se pudo abrir el registro. Vuelve a intentarlo.':error.message;body.replaceChildren(message);setDialog(dialog,true,false);}
    finally{loading=false;}
}
document.addEventListener('click',openDocument);
document.addEventListener('dblclick',openDocument);
document.addEventListener('submit',async event=>{
    const form=event.target.closest('#document-dialog [data-document-form]');if(!form)return;
    event.preventDefault();if(form.dataset.documentBusy==='true')return;if(!form.checkValidity()){form.reportValidity();return;}
    form.dataset.documentBusy='true';setDialog(dialog,true,true);const buttons=[...form.querySelectorAll('button')];buttons.forEach(button=>button.disabled=true);
    try{
        const response=await fetch(form.action,{method:'POST',body:new FormData(form),headers:{'X-Requested-With':'XMLHttpRequest'}});
        if(response.redirected)throw new Error('La sesión o los permisos cambiaron. Vuelve a abrir el listado.');
        if(response.status===422){body.innerHTML=await response.text();dialog.querySelector("h2").textContent=body.querySelector("[data-document-title]")?.dataset.documentTitle??"Registro documental";window.hytInitDates?.(body);body.querySelector('[role=alert]')?.scrollIntoView({block:'nearest'});return;}
        if(!response.ok){const detail=await response.text();throw new Error(!detail.includes('<')&&detail.length<1000?detail:'No se pudo guardar. Revisa el listado antes de reintentar.');}
        const result=await response.json();const destination=new URL(result.redirect,location.href);if(destination.origin!==location.origin)throw new Error('Retorno no válido.');globalThis.hytLoading?.navigation("Cargando listado…");location.assign(destination.href);
    }catch(error){let message=form.querySelector('[data-document-error]');if(!message){message=document.createElement('p');message.dataset.documentError='';message.setAttribute('role','alert');form.prepend(message);}message.textContent=error.message;}
    finally{delete form.dataset.documentBusy;buttons.forEach(button=>button.disabled=false);setDialog(dialog,true,false);}
});