// Firma: FRANCO LARA - 06/11/2026 | Ignora cancelaciones propagadas del selector de archivos; solo el diálogo puede solicitar su cierre.
// Firma: FRANCO LARA - 03/11/2026 | Popup compartido y registro único de eventos aunque se importe con distintas URL de versión.
// Firma: FRANCO LARA - 05/11/2026 | Unifica espera y bloqueo de envíos en el runtime común y protege el cierre de formularios dinámicos ocupados.
// Firma: FRANCO LARA - 10/11/2026 | Conserva la posición del menú por pestaña al navegar.
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
function initializeNavigation() {
    const navigation = document.getElementById('hyt-navigation');
    if (!navigation) return;
    const sidebar = navigation.closest('.hyt-sidebar');
    const toggle = sidebar.querySelector('[data-hyt-menu]');
    const mobile = window.matchMedia('(max-width: 850px)');
    const key = () => mobile.matches ? 'hyt.nav.scroll.mobile' : 'hyt.nav.scroll.desktop';

    function remember() {
        if (!navigation.clientHeight) return;
        try { sessionStorage.setItem(key(), String(navigation.scrollTop)); } catch { }
    }

    function restore() {
        if (!navigation.clientHeight) return;
        try { navigation.scrollTop = Number(sessionStorage.getItem(key())) || 0; } catch { }
    }

    restore();
    window.addEventListener('pagehide', remember);
    toggle?.addEventListener('click', () => {
        remember();
        const open = sidebar.classList.toggle('is-menu-open');
        toggle.setAttribute('aria-expanded', String(open));
        if (open) restore();
    });
}

function initialize() {
if(runtime.initialized)return;
runtime.initialized=true;
initializeNavigation();
document.addEventListener('click',event=>{
    const opener=event.target.closest('[data-hyt-open]');
    if(opener)setDialog(document.getElementById(opener.dataset.hytOpen),true);
    const closer=event.target.closest('[data-hyt-close]');
    if(closer){const dialog=closer.closest('dialog');if(dialog?.getAttribute('aria-busy')!=='true'&&!dialog?.querySelector('[data-hyt-submitting=true]'))setDialog(dialog,false);}
});
// Los envíos nativos y su restauración se controlan en hyt-loading.js.

}
initialize();
