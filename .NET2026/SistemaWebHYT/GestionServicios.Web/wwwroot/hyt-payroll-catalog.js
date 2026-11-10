// Firma: FRANCO LARA - 09/11/2026 | Búsqueda y selección de bancos y previsión en diálogo HYT.
import {setDialog} from './hyt-ui.js';

function searchable(value) {
    return value.normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase();
}

function filter(dialog) {
    const query = searchable(dialog.querySelector('[data-catalog-search]').value.trim());
    const regime = document.getElementById(dialog.dataset.regimeField)?.value;
    let matches = 0;
    for (const row of dialog.querySelectorAll('[data-catalog-row]')) {
        const name = row.querySelector('[data-catalog-select]').dataset.catalogSelect;
        row.hidden = !searchable(name).includes(query) || Boolean(regime && row.dataset.catalogRegime !== regime);
        if (!row.hidden) matches++;
    }
    dialog.querySelector('[data-catalog-empty]').hidden = matches > 0;
}

document.addEventListener('click', event => {
    const opener = event.target.closest('[data-hyt-open]');
    const opened = opener && document.getElementById(opener.dataset.hytOpen);
    if (opened?.matches('[data-payroll-catalog]')) {
        opened.querySelector('[data-catalog-search]').value = '';
        filter(opened);
        return;
    }

    const choice = event.target.closest('[data-catalog-select]');
    if (!choice) return;
    const dialog = choice.closest('[data-payroll-catalog]');
    const field = document.getElementById(dialog.dataset.target);
    field.value = choice.dataset.catalogSelect;
    field.dispatchEvent(new Event('change', {bubbles: true}));
    setDialog(dialog, false);
});

document.addEventListener('input', event => {
    if (event.target.matches('[data-catalog-search]')) filter(event.target.closest('[data-payroll-catalog]'));
});

document.addEventListener('keydown', event => {
    if (event.key === 'Enter' && event.target.matches('[data-catalog-search]')) event.preventDefault();
});
