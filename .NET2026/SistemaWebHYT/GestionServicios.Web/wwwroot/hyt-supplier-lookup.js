// Firma: FRANCO LARA - 09/11/2026 | Ayuda anidada que conserva el borrador del costo.
import {setDialog} from './hyt-ui.js';

async function loadSuppliers(dialog, url) {
    if (!dialog || dialog.getAttribute('aria-busy') === 'true') return;
    const body = dialog.querySelector('[data-supplier-results]');
    const status = dialog.querySelector('[data-supplier-status]');
    status.textContent = 'Cargando proveedores…';
    setDialog(dialog, true, true);
    try {
        const response = await fetch(url, {headers: {'X-Requested-With': 'XMLHttpRequest'}});
        if (!response.ok || response.redirected) throw new Error();
        body.innerHTML = await response.text();
        status.textContent = '';
        body.querySelector('input[type=search]')?.focus();
    } catch {
        status.textContent = 'No se pudo cargar la ayuda de proveedores. Vuelve a intentarlo.';
    } finally {
        setDialog(dialog, true);
    }
}

document.addEventListener('click', event => {
    const opener = event.target.closest('[data-supplier-open]');
    if (opener) {
        const editor = opener.closest('[data-supplier-editor]');
        loadSuppliers(editor.querySelector('[data-supplier-dialog]'), opener.dataset.supplierOpen);
        return;
    }

    const page = event.target.closest('[data-supplier-page]');
    if (page) {
        event.preventDefault();
        loadSuppliers(page.closest('dialog'), page.href);
        return;
    }

    const choice = event.target.closest('[data-supplier-select]');
    if (!choice) return;
    const dialog = choice.closest('dialog');
    if (dialog.getAttribute('aria-busy') === 'true') return;

    const editor = dialog.closest('[data-supplier-editor]');
    editor.querySelector('[name=SupplierId]').value = choice.dataset.supplierSelect;
    editor.querySelector('[data-supplier-label]').textContent = choice.dataset.supplierName;
    setDialog(dialog, false);
});

document.addEventListener('submit', event => {
    const form = event.target.closest('[data-supplier-search]');
    if (!form) return;
    event.preventDefault();
    const url = new URL(form.action, location.href);
    url.search = new URLSearchParams(new FormData(form)).toString();
    loadSuppliers(form.closest('dialog'), url);
});
