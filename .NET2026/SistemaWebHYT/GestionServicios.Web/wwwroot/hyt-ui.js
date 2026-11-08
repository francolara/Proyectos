// Firma: FRANCO LARA - 06/11/2026 | Ignora cancelaciones propagadas del selector de archivos; solo el diálogo puede solicitar su cierre.
// Firma: FRANCO LARA - 03/11/2026 | Popup compartido y registro único de eventos aunque se importe con distintas URL de versión.
// Firma: FRANCO LARA - 05/11/2026 | Unifica espera y bloqueo de envíos en el runtime común y protege el cierre de formularios dinámicos ocupados.
const runtimeKey = Symbol.for('hyt.ui.runtime');
const runtime = globalThis[runtimeKey] ??= {dialogs:new WeakMap(),initialized:false};
const dialogs = runtime.dialogs;
export function setDialog(element, open, busy = false, onClose = null) {
    if (!element) return;
    let state = dialogs.get(element);
    if (!state) {
        state = {busy, onClose, previous:null};
        element.addEventListener('keydown', event => {
            if(event.key==='Escape' && (state.busy||element.querySelector('[data-hyt-submitting=true]'))){event.preventDefault();event.stopPropagation();}
        },true);
        element.addEventListener('cancel', event => {
            if(event.target!==element)return;
            event.preventDefault();
            if(!state.busy&&!element.querySelector('[data-hyt-submitting=true]')) state.onClose ? state.onClose() : setDialog(element,false);
        });
        dialogs.set(element,state);
    }
    state.busy=busy; state.onClose=onClose;
    element.setAttribute('closedby',busy?'none':'closerequest');
    element.setAttribute('aria-busy',String(busy));
    if(open&&!element.open){
        state.previous=document.activeElement;element.showModal();
        (element.querySelector('[autofocus]:not(:disabled)')||element.querySelector('button:not(:disabled),input:not(:disabled)')||element).focus();
    }else if(!open&&element.open){element.close();if(state.previous?.isConnected)state.previous.focus();state.previous=null;}
}
function initialize() {
if(runtime.initialized)return;
runtime.initialized=true;
document.querySelector('[data-hyt-menu]')?.addEventListener('click',event=>{
    const button=event.currentTarget, sidebar=button.closest('.hyt-sidebar');
    button.setAttribute('aria-expanded',String(sidebar.classList.toggle('is-menu-open')));
});
document.addEventListener('click',event=>{
    const opener=event.target.closest('[data-hyt-open]');
    if(opener)setDialog(document.getElementById(opener.dataset.hytOpen),true);
    const closer=event.target.closest('[data-hyt-close]');
    if(closer){const dialog=closer.closest('dialog');if(dialog?.getAttribute('aria-busy')!=='true'&&!dialog?.querySelector('[data-hyt-submitting=true]'))setDialog(dialog,false);}
});
// Los envíos nativos y su restauración se controlan en hyt-loading.js.

}
initialize();
