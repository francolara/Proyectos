// Firma: FRANCO LARA - 05/11/2026 | Filas de enlaces del cliente, índices consecutivos y apertura HTTP/S segura.
(() => {
    const section = document.querySelector('[data-client-links]');
    if (!section) return;
    const rows = section.querySelector('[data-client-link-rows]');
    const status = section.querySelector('[data-client-link-status]');
    function updateOpen(row, index) {
        const input = row.querySelector('[data-link-field="Url"]');
        if (!input) return;
        const anchor = row.querySelector('[data-client-link-open]');
        let url;
        try { const parsed = new URL(input.value.trim()); if (['http:','https:'].includes(parsed.protocol) && !parsed.username && !parsed.password && !/[\u0000-\u001f\u007f]/.test(input.value)) url = parsed.href; } catch { }
        anchor.hidden = !url;
        if (url) anchor.href = url; else anchor.removeAttribute('href');
        anchor.setAttribute('aria-label', `Abrir enlace ${index + 1} en otra pestaña`);
    }
    function refresh() {
        [...rows.children].forEach((row, index) => {
            row.querySelectorAll('[data-link-field]').forEach(input => {
                input.name = `DocumentLinks[${index}].${input.dataset.linkField}`;
                if (input.type !== 'hidden') input.setAttribute('aria-label', `${input.dataset.linkField === 'Url' ? 'URL' : 'Descripción'} del enlace ${index + 1}`);
            });
            updateOpen(row, index);
        });
        section.querySelector('[data-client-link-empty]').hidden = rows.children.length > 0;
    }
    section.querySelector('[data-client-link-add]')?.addEventListener('click', () => {
        if (rows.children.length >= 100) { status.textContent = 'Puedes registrar hasta cien enlaces por cliente.'; return; }
        rows.append(section.querySelector('template').content.cloneNode(true));
        refresh(); rows.lastElementChild.querySelector('[data-link-field="Description"]').focus();
        status.textContent = 'Enlace agregado. Completa la descripción y URL y guarda el cliente.';
    });
    rows.addEventListener('input', event => {
        if (event.target.matches('[data-link-field="Url"]')) refresh();
    });
    rows.addEventListener('click', event => {
        const button = event.target.closest('[data-client-link-remove]');
        if (!button) return;
        const row = button.closest('[data-client-link-row]');
        const next = row.nextElementSibling ?? row.previousElementSibling;
        row.remove(); refresh();
        (next?.querySelector('[data-link-field="Description"]') ?? section.querySelector('[data-client-link-add]')).focus();
        status.textContent = 'Enlace retirado de la lista. Guarda el cliente para confirmar el cambio.';
    });
    refresh();
})();
