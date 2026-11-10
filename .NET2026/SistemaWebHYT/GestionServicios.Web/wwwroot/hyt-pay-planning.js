// Firma: FRANCO LARA - 09/11/2026 | Listado, edición en pop-up, búsqueda de presupuestos y recálculo de ajustes de planificación.
// Firma: FRANCO LARA - 10/11/2026 | Acciones compactas con iconos accesibles en el listado de ajustes. Exige guardar cambios antes de confirmar el pago.
import {setDialog} from './hyt-ui.js';

const money = new Intl.NumberFormat('es-PE', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const amount = value => Number(String(value ?? '').trim().replace(',', '.')) || 0;
const form = document.querySelector('[data-pay-plan]');
const dialog = document.querySelector('#pay-plan-adjustments');

if (form && dialog) {
    const editor = dialog.querySelector('[data-plan-adjustment-editor]');
    const list = dialog.querySelector('[data-plan-adjustment-list]');
    const rows = new Map([...form.querySelectorAll('[data-plan-summary]')]
        .map(row => [row, JSON.parse(row.dataset.planAdjustments)]));
    const field = name => editor.querySelector('[data-adjustment-field="' + name + '"]');
    let worker;
    let editing = -1;
    const confirmButton = document.querySelector('[data-plan-confirm-button]');
    const confirmWarning = document.querySelector('[data-plan-confirm-warning]');

    function markUnsaved() {
        if (!confirmButton) return;
        confirmButton.disabled = true;
        confirmWarning.hidden = false;
    }

    function updateSummary(row) {
        const adjustments = rows.get(row);
        let hours = Number(row.dataset.planLegacyMinutes) / 60;
        let extra = amount(row.dataset.planLegacyExtra);
        let discount = amount(row.dataset.planLegacyDiscount);
        if (adjustments !== null) {
            hours = extra = discount = 0;
            for (const adjustment of adjustments) {
                if (adjustment.Tipo === 'EXTRA') {
                    hours += Number(adjustment.Horas);
                    extra += amount(adjustment.Importe);
                } else {
                    discount += amount(adjustment.Importe);
                }
            }
        }
        row.querySelector('[data-plan-hours-total]').textContent = hours.toLocaleString('es-PE', {maximumFractionDigits: 2});
        row.querySelector('[data-plan-extra-total]').textContent = money.format(extra);
        row.querySelector('[data-plan-discount-total]').textContent = money.format(discount);
        const base = row.querySelector('[data-plan-base]');
        if (base) row.querySelector('[data-plan-net]').textContent = money.format(amount(base.value) + extra - discount);
    }

    function cell(row, value, css) {
        const td = row.insertCell();
        td.textContent = value;
        if (css) td.className = css;
        return td;
    }

    function showList() {
        const adjustments = rows.get(worker);
        list.replaceChildren();
        dialog.querySelector('[data-plan-legacy]').hidden = adjustments !== null ||
            !['planLegacyMinutes', 'planLegacyExtra', 'planLegacyDiscount'].some(key => amount(worker.dataset[key]));
        if (!adjustments?.length) {
            const empty = cell(list.insertRow(), 'No hay ajustes registrados.', 'hyt-grid-message');
            empty.colSpan = editor ? 7 : 6;
            return;
        }

        for (const [index, adjustment] of adjustments.entries()) {
            const row = list.insertRow();
            const extra = adjustment.Tipo === 'EXTRA';
            cell(row, extra ? 'Horas extra' : 'Descuento');
            cell(row, adjustment.Fecha?.split('-').reverse().join('/') || '—');
            cell(row, adjustment.Presupuesto || 'Sin presupuesto', 'hyt-numeric');
            cell(row, adjustment.Concepto || '—');
            cell(row, extra ? adjustment.Horas : '—', 'hyt-numeric');
            cell(row, money.format(amount(adjustment.Importe)), 'hyt-money');
            if (!editor) continue;
            const actions = cell(row, '');
            const group = document.createElement('div');
            group.className = 'hyt-payroll-actions hyt-adjustment-row-actions';
            for (const [action, label, icon] of [
                ['planEdit', 'Editar', 'm16 3 5 5M4 15 16 3a2 2 0 0 1 3 0l2 2a2 2 0 0 1 0 3L9 20l-6 1z'],
                ['planRemove', 'Quitar', 'M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7M14 10v7']
            ]) {
                const button = document.createElement('button');
                button.type = 'button';
                button.className = 'hyt-button-secondary hyt-adjustment-icon';
                button.dataset[action] = index;
                button.title = label;
                button.setAttribute('aria-label', label);
                button.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                    stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${icon}"/></svg>`;
                group.append(button);
            }
            actions.append(group);
        }
    }

    function saveDraft(adjustments) {
        markUnsaved();
        rows.set(worker, adjustments);
        worker.querySelector('[data-plan-detailed]').value = 'true';
        const inputs = worker.querySelector('[data-plan-fields]');
        inputs.replaceChildren();
        for (const [index, adjustment] of adjustments.entries()) {
            for (const [key, value] of Object.entries(adjustment)) {
                const input = document.createElement('input');
                input.type = 'hidden';
                input.name = `Lines[${worker.dataset.planSummary}].Ajustes[${index}].${key}`;
                input.value = value ?? '';
                inputs.append(input);
            }
        }
        updateSummary(worker);
        showList();
    }

    function syncEditor() {
        const extra = field('Tipo').value === 'EXTRA';
        editor.querySelector('[data-plan-hours]').hidden = !extra;
        field('Horas').min = extra ? '1' : '0';
        if (!extra) field('Horas').value = '0';
        if (extra && field('Horas').value === '0') field('Horas').value = '1';
        field('Concepto').required = !extra || !field('Presupuesto').value.trim();
        field('Importe').setCustomValidity('');
    }

    function openEditor(index, type) {
        editing = index;
        editor.reset();
        const adjustment = rows.get(worker)?.[index];
        if (adjustment) {
            for (const [key, value] of Object.entries(adjustment)) field(key).value = value ?? '';
        } else {
            field('Tipo').value = type;
        }
        syncEditor();
        editor.querySelector('[data-plan-editor-title]').textContent = adjustment ? 'Editar ajuste' : 'Nuevo ajuste';
        editor.hidden = false;
        field('Fecha').focus();
    }

    form.addEventListener('input', event => {
        if (event.target.matches('[data-plan-base]')) {
            markUnsaved();
            updateSummary(event.target.closest('[data-plan-summary]'));
        }
    });
    form.addEventListener('click', event => {
        const button = event.target.closest('[data-plan-open]');
        if (!button) return;
        worker = button.closest('[data-plan-summary]');
        dialog.querySelector('[data-plan-worker-name-label]').textContent = worker.dataset.planWorkerName;
        if (editor) editor.hidden = true;
        showList();
        setDialog(dialog, true);
    });
    dialog.addEventListener('click', event => {
        const add = event.target.closest('[data-plan-add]');
        const edit = event.target.closest('[data-plan-edit]');
        const remove = event.target.closest('[data-plan-remove]');
        if (add) openEditor(-1, add.dataset.planAdd);
        if (edit) openEditor(Number(edit.dataset.planEdit));
        if (remove) {
            editor.hidden = true;
            saveDraft(rows.get(worker).filter((_, index) => index !== Number(remove.dataset.planRemove)));
            dialog.querySelector('[data-plan-add]')?.focus();
        }
        if (event.target.closest('[data-plan-cancel-edit]')) {
            editor.hidden = true;
            dialog.querySelector('[data-plan-add]')?.focus();
        }
    });
    editor?.addEventListener('input', syncEditor);
    editor?.addEventListener('submit', event => {
        event.preventDefault();
        const type = field('Tipo').value;
        const raw = field('Importe').value.trim();
        if (!/^\d+(?:[.,]\d{1,2})?$/.test(raw) || (type === 'DESCUENTO' && amount(raw) <= 0)) {
            field('Importe').setCustomValidity('Indica un importe válido con hasta dos decimales.');
            field('Importe').reportValidity();
            return;
        }
        const adjustments = [...(rows.get(worker) ?? [])];
        const value = {
            Tipo: type, Fecha: field('Fecha').value, Presupuesto: field('Presupuesto').value.trim(),
            Concepto: field('Concepto').value.trim(), Horas: Number(field('Horas').value), Importe: raw.replace(',', '.')
        };
        if (editing < 0) adjustments.push(value);
        else adjustments[editing] = value;
        saveDraft(adjustments);
        editor.hidden = true;
        dialog.querySelector('[data-plan-add]')?.focus();
    });
}
const budgetDialog = document.querySelector('#pay-plan-budget');
if (budgetDialog) {
    const opener = dialog.querySelector('[data-plan-budget-open]');
    const budgetField = dialog.querySelector('[data-adjustment-field="Presupuesto"]');
    const searchForm = budgetDialog.querySelector('[data-plan-budget-search]');
    const searchInput = searchForm.elements.namedItem('search');
    const status = budgetDialog.querySelector('[data-plan-budget-status]');
    const results = budgetDialog.querySelector('[data-plan-budget-results]');
    let search = '';
    let pending;

    async function loadBudgets(page) {
        pending?.abort();
        const request = new AbortController();
        pending = request;
        const url = new URL(searchForm.action);
        url.searchParams.set('search', search);
        url.searchParams.set('page', page);
        results.replaceChildren();
        status.hidden = false;
        status.textContent = 'Buscando presupuestos…';
        try {
            const response = await fetch(url, {signal: request.signal});
            if (!response.ok || response.redirected) throw new Error();
            const html = await response.text();
            if (pending !== request || !budgetDialog.open) return;
            results.innerHTML = html;
            status.hidden = true;
            if (page > 1) results.querySelector('[role="region"]').focus();
        } catch {
            if (pending !== request || request.signal.aborted || !budgetDialog.open) return;
            status.textContent = 'No se pudieron cargar los presupuestos. Vuelve a pulsar Buscar.';
        }
    }

    opener.addEventListener('click', () => {
        search = searchInput.value = budgetField.value.trim();
        setDialog(budgetDialog, true);
        loadBudgets(1);
    });
    searchForm.addEventListener('submit', event => {
        event.preventDefault();
        search = searchInput.value.trim();
        loadBudgets(1);
    });
    results.addEventListener('click', event => {
        const selection = event.target.closest('[data-plan-budget-number]');
        const page = event.target.closest('[data-plan-budget-page]');
        if (page) loadBudgets(Number(page.dataset.planBudgetPage));
        if (selection) {
            budgetField.value = selection.dataset.planBudgetNumber;
            budgetField.dispatchEvent(new Event('input', {bubbles: true}));
            setDialog(budgetDialog, false);
        }
    });
    budgetDialog.addEventListener('keydown', event => {
        if (event.key !== 'Escape') return;
        event.preventDefault();
        setDialog(budgetDialog, false);
    });
    budgetDialog.addEventListener('close', () => {
        if (!budgetDialog.open) pending?.abort();
    });
}
