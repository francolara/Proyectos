// Firma: FRANCO LARA - 09/11/2026 | Listado, edición en pop-up, búsqueda de presupuestos y recálculo de ajustes de planificación.
// Firma: FRANCO LARA - 10/11/2026 | Acciones compactas con iconos accesibles en el listado de ajustes. Exige guardar cambios antes de confirmar el pago y calcula abonos y cargos por concepto y cambia entre selector de fecha y quincena. Exige guardar los cambios antes de exportar el TXT BCP.
// Firma: FRANCO LARA - 11/11/2026 | Evita editar períodos con filtros sin cargar, sincroniza el calendario y actualiza los totales y el detalle de abonos y cargos y excluye de los totales las filas desmarcadas.
import {setDialog} from './hyt-ui.js';

const money = new Intl.NumberFormat('es-PE', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const amount = value => Number(String(value ?? '').trim().replace(',', '.')) || 0;
const periodFilter = document.querySelector('[data-plan-period]');
if (periodFilter) {
    const frequency = periodFilter.elements.namedItem('frequency');
    const date = periodFilter.elements.namedItem('from');
    const month = periodFilter.elements.namedItem('month');
    const half = periodFilter.elements.namedItem('half');
    const pending = document.querySelector('[data-plan-period-pending]');
    const results = document.querySelectorAll('[data-plan-period-results]');
    const syncPeriod = () => {
        const isBiweekly = frequency.value === 'QUINCENAL';
        const selectedFrom = isBiweekly ? (month.value ? month.value + (half.value === '2' ? '-16' : '-01') : '') : date.value;
        const changed = frequency.value !== periodFilter.dataset.loadedFrequency || selectedFrom !== periodFilter.dataset.loadedFrom;
        month.disabled = half.disabled = !isBiweekly;
        periodFilter.querySelector('[data-plan-date-field]').hidden = isBiweekly;
        for (const field of periodFilter.querySelectorAll('[data-plan-half-field]')) field.hidden = !isBiweekly;
        pending.hidden = !changed;
        for (const result of results) result.hidden = changed;
    };
    frequency.addEventListener('change', () => {
        const isBiweekly = frequency.value === 'QUINCENAL';
        if (isBiweekly && date.value) {
            month.value = date.value.slice(0, 7);
            half.value = Number(date.value.slice(8)) <= 15 ? '1' : '2';
        }
    });
    periodFilter.addEventListener('change', syncPeriod);
    syncPeriod();
}

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
    let percentage = null;
    let percentageBase = null;
    const conceptSelect = editor ? field('ConceptoId') : null;
    const baseOf = row => amount(row.querySelector('[data-plan-base]')?.value ?? row.dataset.planBase);
    const summaries = new Map([...rows.keys()].map(row => [row, {
        base: amount(row.dataset.planBase),
        abonos: amount(row.dataset.planLegacyExtra),
        cargos: amount(row.dataset.planLegacyDiscount)
    }]));
    const footer = [...form.querySelectorAll('[data-plan-footer]')];
    const calculatedAmount = (row, adjustment) => adjustment.Porcentaje == null
        ? amount(adjustment.Importe)
        : Math.round(((adjustment.BasePorcentaje ?? baseOf(row)) * Number(adjustment.Porcentaje)) + 1e-9) / 100;
    const confirmButton = document.querySelector('[data-plan-confirm-button]');
    const printLink = document.querySelector('[data-plan-print]');
    const exportButtons = document.querySelectorAll('[data-plan-export]');
    const confirmWarning = document.querySelector('[data-plan-confirm-warning]');

    function markUnsaved() {
        if (printLink) printLink.hidden = true;
        if (confirmButton) confirmButton.disabled = true;
        for (const button of exportButtons) button.disabled = true;
        if (confirmWarning) confirmWarning.hidden = false;
    }

    function updateSummary(row) {
        const adjustments = rows.get(row);
        let extra = amount(row.dataset.planLegacyExtra);
        let discount = amount(row.dataset.planLegacyDiscount);
        if (adjustments !== null) {
            extra = discount = 0;
            for (const adjustment of adjustments) {
                const value = calculatedAmount(row, adjustment);
                if (adjustment.Tipo === 'DESCUENTO') discount += value;
                else extra += value;
            }
        }
        row.querySelector('[data-plan-total="abonos"]').textContent = money.format(extra);
        row.querySelector('[data-plan-total="cargos"]').textContent = money.format(discount);
        const base = baseOf(row);
        row.querySelector('[data-plan-net]').textContent = money.format(base + extra - discount);
        if (adjustments !== null) updateInlineDetails(row, adjustments);

        summaries.set(row, {base, abonos: extra, cargos: discount});
        updateTotals();
    }

    function updateTotals() {
        const totals = {base: 0, abonos: 0, cargos: 0, neto: 0};
        for (const [row, summary] of summaries) {
            if (row.dataset.planIncluded === 'false') continue;
            totals.base += summary.base;
            totals.abonos += summary.abonos;
            totals.cargos += summary.cargos;
        }
        totals.neto = totals.base + totals.abonos - totals.cargos;
        for (const cell of footer) cell.textContent = money.format(totals[cell.dataset.planFooter]);
    }

    function updateInlineDetails(row, adjustments) {
        for (const toggle of row.querySelectorAll('[data-hyt-row-toggle]')) {
            const list = document.getElementById(toggle.getAttribute('aria-controls')).querySelector('[data-plan-inline-list]');
            const charge = list.dataset.planInlineList === 'cargos';
            const items = adjustments.filter(adjustment => (adjustment.Tipo === 'DESCUENTO') === charge);
            list.replaceChildren();
            for (const adjustment of items) {
                const item = document.createElement('li');
                const date = document.createElement('span');
                const concept = document.createElement('span');
                const value = document.createElement('strong');
                date.textContent = adjustment.Fecha?.split('-').reverse().join('/') || '—';
                concept.textContent = adjustment.Concepto || (adjustment.Tipo === 'EXTRA' ? 'Horas extra' : '—');
                value.textContent = 'S/ ' + money.format(calculatedAmount(row, adjustment));
                const budget = document.createElement('span');
                budget.className = 'hyt-muted';
                budget.textContent = adjustment.Presupuesto ? 'Presupuesto ' + adjustment.Presupuesto : '';
                value.className = 'hyt-money';
                item.append(date, concept, budget, value);
                list.append(item);
            }
            if (!items.length) {
                const empty = document.createElement('li');
                empty.textContent = 'Sin detalle registrado';
                list.append(empty);
            }
        }
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
            cell(row, extra ? 'Horas extra' : adjustment.Tipo === 'ABONO' ? 'Abono' : 'Cargo');
            cell(row, adjustment.Fecha?.split('-').reverse().join('/') || '—');
            cell(row, adjustment.Presupuesto || 'Sin presupuesto', 'hyt-numeric');
            cell(row, adjustment.Concepto || '—');
            cell(row, extra ? adjustment.Horas : '—', 'hyt-numeric');
            cell(row, money.format(calculatedAmount(worker, adjustment)), 'hyt-money');
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
        const type = field('Tipo').value;
        const extra = type === 'EXTRA';
        editor.querySelector('[data-plan-hours]').hidden = !extra;
        editor.querySelector('[data-plan-concept]').hidden = extra;
        for (const option of conceptSelect.options) {
            option.hidden = option.disabled = !!option.dataset.kind && option.dataset.kind !== type;
        }
        conceptSelect.required = type === 'ABONO';
        const hasConcept = !extra && !!conceptSelect.value;
        const showRate = hasConcept && (percentage !== null ||
            conceptSelect.selectedOptions[0]?.dataset.calculation === 'PORCENTAJE');
        editor.querySelector('[data-plan-rate]').hidden = !showRate;
        field('Porcentaje').disabled = !showRate;
        const rateLabel = editor.querySelector('[data-plan-rate-label]');
        if (rateLabel) rateLabel.textContent = percentageBase === null ? 'Porcentaje sobre la base (opcional)'
            : 'Porcentaje sobre el sueldo mensual de S/ ' + money.format(percentageBase);
        editor.querySelector('[data-plan-description]').hidden = hasConcept;
        field('Horas').min = extra ? '1' : '0';
        if (!extra) field('Horas').value = '0';
        if (extra && field('Horas').value === '0') field('Horas').value = '1';
        field('Concepto').required = !hasConcept && (type === 'DESCUENTO' || extra && !field('Presupuesto').value.trim());
        field('Importe').setCustomValidity('');
    }

    function openEditor(index, type) {
        editing = index;
        editor.reset();
        conceptSelect.querySelector('[data-plan-previous-option]')?.remove();
        const adjustment = rows.get(worker)?.[index];
        percentage = adjustment?.Porcentaje ?? null;
        percentageBase = adjustment?.BasePorcentaje ?? null;
        if (adjustment) {
            for (const [key, value] of Object.entries(adjustment)) {
                const control = field(key);
                if (control && key !== 'ConceptoId') control.value = value ?? '';
            }
            if (adjustment.ConceptoId != null) {
                const id = String(adjustment.ConceptoId);
                if (![...conceptSelect.options].some(option => option.value === id && option.dataset.kind === adjustment.Tipo)) {
                    const previous = new Option(adjustment.Concepto, id);
                    previous.dataset.kind = adjustment.Tipo;
                    if (percentage !== null || percentageBase !== null) previous.dataset.calculation = 'PORCENTAJE';
                    previous.dataset.planPreviousOption = '';
                    conceptSelect.prepend(previous);
                }
                conceptSelect.value = id;
            }
            field('Importe').value = calculatedAmount(worker, adjustment).toFixed(2);
        } else {
            field('Tipo').value = type;
        }
        field('Tipo').querySelector('option[value="EXTRA"]').disabled = adjustment?.Tipo !== 'EXTRA';
        syncEditor();
        editor.querySelector('[data-plan-editor-title]').textContent = adjustment ? 'Editar ajuste' : 'Nuevo ajuste';
        editor.hidden = false;
        const dateField = field('Fecha');
        dateField.dispatchEvent(new Event('change', { bubbles: true }));
        (dateField.closest('.app-date-picker')?.querySelector('.app-date-picker-display') ?? dateField).focus();
    }

    form.addEventListener('input', event => {
        if (event.target.matches('[data-plan-base]')) {
            markUnsaved();
            updateSummary(event.target.closest('[data-plan-summary]'));
        }
    });
    form.addEventListener('change', event => {
        const checkbox = event.target.closest('[data-plan-include]');
        if (!checkbox) return;
        checkbox.closest('[data-plan-summary]').dataset.planIncluded = String(checkbox.checked);
        markUnsaved();
        updateTotals();
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
    editor?.addEventListener('input', event => {
        if (event.target === field('Importe')) {
            percentage = null;
            field('Porcentaje').value = '';
        }
        if (event.target === field('Porcentaje')) {
            percentage = field('Porcentaje').value === '' ? null : field('Porcentaje').valueAsNumber;
            if (Number.isFinite(percentage)) {
                field('Importe').value = calculatedAmount(worker, { Porcentaje: percentage, BasePorcentaje: percentageBase }).toFixed(2);
            }
        }
        if (event.target === field('Tipo')) {
            conceptSelect.value = '';
            field('Concepto').value = '';
            field('Importe').value = '0.00';
            field('Porcentaje').value = '';
            percentage = percentageBase = null;
        }
        if (event.target === conceptSelect) {
            const option = conceptSelect.selectedOptions[0];
            const defaultValue = option?.dataset.value;
            percentageBase = option?.dataset.monthlyBasis === 'true' && option.dataset.calculation === 'PORCENTAJE'
                ? amount(worker.dataset.planMonthlySalary) : null;
            percentage = option?.dataset.calculation === 'PORCENTAJE' && defaultValue
                ? Number(defaultValue) : null;
            field('Porcentaje').value = percentage ?? '';
            field('Importe').value = (percentage === null ? amount(defaultValue)
                : calculatedAmount(worker, { Porcentaje: percentage, BasePorcentaje: percentageBase })).toFixed(2);
        }
        syncEditor();
    });
    editor?.addEventListener('submit', event => {
        event.preventDefault();
        const type = field('Tipo').value;
        const raw = field('Importe').value.trim();
        const conceptId = type !== 'EXTRA' && conceptSelect.value ? Number(conceptSelect.value) : null;
        if (!/^\d+(?:[.,]\d{1,2})?$/.test(raw) || (type === 'DESCUENTO' && !conceptId && amount(raw) <= 0)) {
            field('Importe').setCustomValidity('Indica un importe válido con hasta dos decimales.');
            field('Importe').reportValidity();
            return;
        }
        const adjustments = [...(rows.get(worker) ?? [])];
        const value = {
            Tipo: type, Fecha: field('Fecha').value, Presupuesto: field('Presupuesto').value.trim(),
            Concepto: conceptId ? conceptSelect.selectedOptions[0].text : field('Concepto').value.trim(),
            Horas: Number(field('Horas').value), Importe: raw.replace(',', '.'),
            ConceptoId: conceptId, Porcentaje: percentage, BasePorcentaje: percentageBase
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
