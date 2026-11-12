// Firma: FRANCO LARA - 10/11/2026 | Ayudas de terceros, búsqueda de presupuestos y cálculo automático de importes, subtotal, IGV, total y detracción con moneda de depósito en compras.
import { setDialog } from "./hyt-ui.js";
function searchOptions(form, inputSelector, selectSelector, statusSelector, endpoint, idKey, label) {
    if (!form) return;
    const input = form.querySelector(inputSelector);
    const select = form.querySelector(selectSelector);
    const status = form.querySelector(statusSelector);
    let timer, request = 0;
    input.addEventListener('input', () => {
        clearTimeout(timer);
        const current = ++request;
        const query = input.value.trim();
        select.replaceChildren(new Option('Selecciona un registro', ''));
        if (query.length < 2) { status.textContent = 'Escribe al menos dos caracteres.'; return; }
        timer = setTimeout(async () => {
            status.textContent = 'Buscando…';
            try {
                const url = new URL(endpoint, location.origin);
                url.searchParams.set('buscar', query);
                const response = await fetch(url, { headers: { Accept: 'application/json' } });
                if (!response.ok) throw new Error();
                const rows = await response.json();
                if (request !== current) return;
                for (const row of rows) select.add(new Option(label(row), row[idKey]));
                status.textContent = rows.length ? 'Selecciona un resultado.' : 'No se encontraron registros.';
            } catch {
                if (request === current) status.textContent = 'No se pudo completar la búsqueda. Vuelve a escribir para reintentar.';
            }
        }, 300);
    });
}
const costs = document.querySelector('[data-commercial-cost-form]');
searchOptions(costs, '[data-commercial-budget-search]', '[data-commercial-budgets]', '[data-commercial-budget-status]',
    costs?.dataset.budgetsUrl, 'PresupuestoId', row => row.NumeroPresupuesto + ' · ' + row.Descripcion);
// Firma: FRANCO LARA - 11/11/2026 | Precarga la detracción por material y abre las ayudas de órdenes de compra y de presupuestos del cliente y actualiza los correlativos al modificar partidas y calcula el vencimiento de compras y ventas con el plazo elegido y busca facturas emitidas para notas de crédito.
const form = document.querySelector('[data-commercial-form]');
if (form) {
    const body = form.querySelector('[data-commercial-lines]');
    const currency = form.querySelector('[name="Input.Moneda"]');
    const date = form.querySelector('[name="Input.Fecha"]');
    const type = form.querySelector('[name="Input.Tipo"]');
    const payment = form.querySelector('[data-commercial-payment]');
    if (payment) {
        const method = payment.querySelector('[name="Input.FormaPago"]');
        const term = payment.querySelector('[name="Input.PlazoPagoDias"]');
        const dueDate = payment.querySelector('#commercial-due-date');
        const updateDueDate = () => {
            const credit = method.value === 'Crédito';
            term.disabled = !credit;
            term.required = credit;
            if (!date.value || !date.validity.valid || (credit && !term.value)) {
                dueDate.value = '';
                return;
            }
            const days = credit ? Number(term.value) : 0;
            const due = new Date(date.valueAsNumber + days * 86400000);
            if (due.getUTCFullYear() > 9999) {
                dueDate.value = '';
                return;
            }
            const [year, month, day] = due.toISOString().slice(0, 10).split('-');
            dueDate.value = day + '/' + month + '/' + year;
        };
        payment.addEventListener('change', updateDueDate);
        updateDueDate();
    }
    const withholdingUrl = form.dataset.commercialWithholdingUrl;
    const materialStatus = form.querySelector('[data-commercial-material-withholding-status]');
    const withholdingFields = form.querySelector('[data-commercial-withholding-fields]');
    const lookupButton = form.querySelector('[data-hyt-open="commercial-withholding-lookup"]');
    const clearWithholding = form.querySelector('[data-commercial-clear-withholding]');
    let materialWithholding;
    let materialRequest;
    let materialTimer;
    let requestedCodes;
    const formatter = new Intl.NumberFormat('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const percent = new Intl.NumberFormat('es-PE', { style: 'percent', maximumFractionDigits: 4 });
    const subtotalOutput = form.querySelector('[data-commercial-lines-total]');
    const taxLabel = form.querySelector('[data-commercial-tax-label]');
    const taxOutput = form.querySelector('[data-commercial-tax]');
    const totalOutput = form.querySelector('[data-commercial-total]');
    const taxStatus = form.querySelector('[data-commercial-tax-status]');
    const withholdingCode = form.querySelector('#commercial-withholding-code');
    const withholdingName = form.querySelector('#commercial-withholding-name');
    const withholdingAmount = form.querySelector('#commercial-withholding-amount');
    const withholdingRate = form.querySelector('#commercial-withholding-rate');
    const withholdingCurrency = form.querySelector('[name="Input.MonedaDetraccion"]');
    const foreignDeposit = form.querySelector('[name="Input.ImporteDetraccion"]');
    const foreignDepositField = form.querySelector('[data-commercial-foreign-withholding]');
    const equivalentField = form.querySelector('[data-commercial-withholding-equivalent]');
    const equivalent = form.querySelector('#commercial-withholding-equivalent');
    const withholdingOutput = form.querySelector('[data-commercial-withholding]');
    const supplierBalanceOutput = form.querySelector('[data-commercial-supplier-balance]');
    const withholdingOptions = new Map(
        [...form.querySelectorAll('#commercial-withholding-lookup [data-catalog-select]')]
            .map(option => [option.dataset.catalogSelect, option.dataset]));
    let rate = form.dataset.commercialTaxRate ? Number(form.dataset.commercialTaxRate) : null;
    let taxRequest;
    const cents = value => Math.round((value + Number.EPSILON) * 100);
    const money = amount => currency.value + ' ' + formatter.format(amount / 100);

    const refreshTotal = () => {
        let subtotal = 0;
        for (const row of body.rows) {
            const quantity = row.querySelector('[name$=".Cantidad"]').valueAsNumber;
            const price = row.querySelector('[name$=".Precio"]').valueAsNumber;
            const amount = cents((quantity || 0) * (price || 0));
            row.querySelector('[name$=".Importe"]').value = Number.isFinite(amount) ? (amount / 100).toFixed(2) : '';
            subtotal += amount;
        }
        subtotalOutput.textContent = Number.isFinite(subtotal) ? money(subtotal) : '—';
        taxLabel.textContent = rate === null ? 'IGV' : 'IGV (' + percent.format(rate) + ')';
        const tax = rate === null ? NaN : cents(subtotal * rate / 100);
        taxOutput.textContent = Number.isFinite(tax) ? money(tax) : '—';
        const total = subtotal + tax;
        totalOutput.textContent = Number.isFinite(total) ? money(total) : '—';
        if (!withholdingCode) return;
        const salesInvoice = !!withholdingUrl && type.value === '01';
        if (withholdingUrl) withholdingFields.hidden = !salesInvoice;
        const withholding = withholdingUrl
            ? (salesInvoice ? (materialWithholding ?? withholdingOptions.get(withholdingCode.value)) : null)
            : withholdingOptions.get(withholdingCode.value);
        withholdingName.value = withholding?.catalogLabel ?? '';
        if (withholdingRate) withholdingRate.value = withholding ? percent.format(Number(withholding.withholdingRate) / 100) : '';
        lookupButton.disabled = salesInvoice && !!materialWithholding;
        clearWithholding.disabled = salesInvoice && !!materialWithholding;
        if (withholdingCurrency) withholdingCurrency.disabled = !withholding;
        const depositCurrency = withholdingCurrency?.value || currency.value;
        const differentCurrency = !!withholding && depositCurrency !== currency.value;
        if (foreignDeposit) {
            foreignDepositField.hidden = equivalentField.hidden = !differentCurrency;
            foreignDeposit.required = differentCurrency;
            foreignDeposit.disabled = !differentCurrency;
        }

        const deposit = Math.round(total * Number(withholding?.withholdingRate ?? 0) / 100);
        withholdingAmount.value = Number.isFinite(deposit) ? money(deposit) : '';
        if (equivalent) equivalent.value = withholdingAmount.value;
        withholdingOutput.textContent = differentCurrency
            ? (Number.isFinite(foreignDeposit.valueAsNumber) ? depositCurrency + ' ' + formatter.format(foreignDeposit.valueAsNumber) : '—')
            : (withholdingAmount.value || '—');
        supplierBalanceOutput.textContent = Number.isFinite(deposit) ? money(total - deposit) : '—';
    };

    const refreshMaterialWithholding = () => {
        if (!withholdingUrl) return;
        const codes = type.value === '01'
            ? [...new Set([...body.rows].map(row => row.querySelector('[name$=".Codigo"]').value.trim()).filter(Boolean))]
            : [];
        const serialized = JSON.stringify(codes);
        if (requestedCodes === serialized) return;
        requestedCodes = serialized;
        materialRequest?.abort();
        clearTimeout(materialTimer);
        materialWithholding = null;
        materialStatus.hidden = codes.length === 0;
        refreshTotal();
        if (codes.length === 0) return;
        materialStatus.textContent = 'Consultando la detracción de las partidas…';
        const request = new AbortController();
        materialRequest = request;
        materialTimer = setTimeout(async () => {
            try {
                const response = await fetch(withholdingUrl, {
                    method: 'POST', signal: request.signal, body: serialized,
                    headers: { 'Content-Type': 'application/json',
                        RequestVerificationToken: form.querySelector('[name="__RequestVerificationToken"]').value }
                });
                if (!response.ok) throw new Error();
                const result = await response.json();
                if (request.signal.aborted) return;
                materialWithholding = result ? {
                    catalogLabel: withholdingOptions.get(result.Codigo)?.catalogLabel ?? result.Codigo,
                    withholdingRate: result.Porcentaje
                } : null;
                materialStatus.hidden = !materialWithholding;
                materialStatus.textContent = materialWithholding ? 'Porcentaje aplicado según las partidas.' : '';
                refreshTotal();
            } catch {
                if (request.signal.aborted) return;
                materialStatus.textContent = 'La detracción de las partidas se calculará al guardar.';
            }
        }, 200);
    };

    if (withholdingCode) {
        withholdingCode.addEventListener('change', refreshTotal);
        withholdingCurrency?.addEventListener('change', () => {
            foreignDeposit.value = '';
            refreshTotal();
        });
        foreignDeposit?.addEventListener('input', refreshTotal);
        clearWithholding.addEventListener('click', () => {
            withholdingCode.value = '';
            if (withholdingCurrency) withholdingCurrency.value = '';
            if (foreignDeposit) foreignDeposit.value = '';
            refreshTotal();
        });
    }
    body.addEventListener('input', event => {
        if (event.target.matches('[name$=".Codigo"]')) refreshMaterialWithholding();
        refreshTotal();
    });
    currency.addEventListener('change', refreshTotal);
    date.addEventListener('change', async () => {
        taxRequest?.abort();
        rate = null;
        refreshTotal();
        if (!date.value) {
            taxStatus.textContent = 'Selecciona la fecha de emisión.';
            return;
        }
        const request = new AbortController();
        taxRequest = request;
        taxStatus.textContent = 'Consultando la tasa de IGV…';
        try {
            const url = new URL(form.dataset.commercialTaxUrl, location.origin);
            url.searchParams.set('fecha', date.value);
            const response = await fetch(url, { signal: request.signal });
            if (!response.ok) throw new Error();
            const result = await response.json();
            if (request.signal.aborted) return;
            rate = result.tasa;
            taxStatus.textContent = rate === null ? 'No hay una tasa de IGV configurada para esta fecha.' : '';
            refreshTotal();
        } catch {
            if (!request.signal.aborted) taxStatus.textContent = 'No se pudo consultar la tasa de IGV.';
        }
    });
    if (rate === null) taxStatus.textContent = 'No hay una tasa de IGV configurada para esta fecha.';
    refreshTotal();
    const renumber = () => [...body.rows].forEach((row, index) => {
        row.querySelector('[data-commercial-item-number]').textContent = index + 1;
        row.querySelectorAll('[name]').forEach(input => {
            input.name = input.name.replace(/Detalle\[(?:\d+|__index__)\]/, 'Detalle[' + index + ']');
        });
    });
    form.querySelector('[data-commercial-add]').addEventListener('click', () => {
        if (body.rows.length >= 500) return;
        body.append(form.querySelector('[data-commercial-line-template]').content.cloneNode(true));
        renumber();
        refreshTotal();
        body.lastElementChild.querySelector('[name$=".Descripcion"]').focus();
    });
    body.addEventListener('click', event => {
        const button = event.target.closest('[data-commercial-remove]');
        if (button && body.rows.length > 1) {
            button.closest('tr').remove();
            renumber();
            refreshMaterialWithholding();
            refreshTotal();
        }
    });
    const note = form.querySelector('[data-commercial-note]');
    const showNote = () => {
        note.hidden = type.value !== '07';
        note.querySelectorAll('[data-note-required]').forEach(field => field.required = !note.hidden);
    };
    type.addEventListener('change', () => {
        showNote();
        refreshMaterialWithholding();
        refreshTotal();
    });
    showNote();
    refreshMaterialWithholding();
    form.addEventListener('submit', renumber);
}

const lookup = document.querySelector('[data-commercial-parties-url]');
if (lookup && form) {
    const search = lookup.querySelector('[data-commercial-party-search]');
    const query = search.querySelector('input');
    const results = lookup.querySelector('[data-commercial-party-results]');
    const status = lookup.querySelector('[data-commercial-party-status]');
    const sales = lookup.dataset.sales === 'true';
    let currentRequest;
    lookup.addEventListener('close', () => currentRequest?.abort());
    const loadParties = async () => {
        const term = query.value.trim();
        currentRequest?.abort();
        const request = new AbortController();
        currentRequest = request;
        results.replaceChildren();
        status.textContent = 'Buscando…';
        try {
            const url = new URL(lookup.dataset.commercialPartiesUrl, location.origin);
            url.searchParams.set('buscar', term);
            const response = await fetch(url, { signal: request.signal });
            if (!response.ok) throw new Error();
            const parties = await response.json();
            if (request.signal.aborted) return;
            for (const party of parties) {
                const row = results.insertRow();
                row.insertCell().textContent = party.DocumentoTercero;
                row.insertCell().textContent = party.NombreTercero;
                const button = document.createElement('button');
                button.type = 'button';
                button.className = 'hyt-button-secondary hyt-action-button';
                button.textContent = 'Seleccionar';
                button.addEventListener('click', () => {
                    if (sales) {
                        const page = new URL(lookup.dataset.commercialNewUrl, location.origin);
                        page.searchParams.set('clienteId', party.ClienteId);
                        location.assign(page);
                        return;
                    }
                    form.querySelector('[name="Input.DocumentoTercero"]').value = party.DocumentoTercero;
                    form.querySelector('[name="Input.NombreTercero"]').value = party.NombreTercero;
                    setDialog(lookup, false);
                });
                row.insertCell().append(button);
            }
            status.textContent = parties.length ? 'Selecciona un resultado.' : 'No se encontraron registros.';
            if (parties.length === 30) status.textContent = 'Se muestran 30 resultados. Escribe un nombre o documento para acotar la búsqueda.';
        } catch {
            if (!request.signal.aborted) status.textContent = 'No se pudo completar la búsqueda. Inténtalo nuevamente.';
        }
    };
    document.querySelector('[data-commercial-party-open]').addEventListener('click', () => {
        setDialog(lookup, true);
        void loadParties();
    });
    search.addEventListener('submit', event => {
        event.preventDefault();
        void loadParties();
    });
}

const orderPicker = document.querySelector('[data-commercial-order-picker]');
if (orderPicker?.dataset.orderUrl) {
    const dialog = orderPicker.querySelector('dialog');
    const query = dialog.querySelector('input[type="search"]');
    const results = dialog.querySelector('[data-order-results]');
    const status = dialog.querySelector('[data-order-status]');
    const id = orderPicker.querySelector('[data-order-id]');
    const label = orderPicker.querySelector('[data-order-label]');
    const clear = orderPicker.querySelector('[data-order-clear]');
    let pending;

    async function searchOrders() {
        pending?.abort();
        const request = new AbortController();
        pending = request;
        results.replaceChildren();
        status.textContent = 'Buscando órdenes…';
        try {
            const url = new URL(orderPicker.dataset.orderUrl, location.origin);
            url.searchParams.set('buscar', query.value.trim());
            const response = await fetch(url, { signal: request.signal });
            if (!response.ok || response.redirected) throw new Error();
            const orders = await response.json();
            if (request.signal.aborted) return;
            for (const order of orders) {
                const row = results.insertRow();
                row.insertCell().textContent = order.NumeroOCCliente;
                row.insertCell().textContent = order.NumeroPresupuesto;
                row.insertCell().textContent = order.NumeroConfirmacion ?? '—';
                const button = document.createElement('button');
                button.type = 'button';
                button.className = 'hyt-button-secondary hyt-action-button';
                button.textContent = 'Seleccionar';
                button.addEventListener('click', () => {
                    id.value = order.OrdenCompraId;
                    label.value = order.NumeroOCCliente;
                    clear.disabled = false;
                    setDialog(dialog, false);
                });
                row.insertCell().append(button);
            }
            status.textContent = orders.length === 30 ? 'Se muestran 30 órdenes. Usa la búsqueda para acotar los resultados.'
                : orders.length ? '' : 'No se encontraron órdenes para este cliente.';
        } catch {
            if (!request.signal.aborted) status.textContent = 'No se pudo consultar la lista de órdenes.';
        }
    }

    orderPicker.querySelector('[data-order-open]').addEventListener('click', () => {
        setDialog(dialog, true);
        void searchOrders();
    });
    clear.addEventListener('click', () => {
        id.value = '';
        label.value = '';
        clear.disabled = true;
    });
    dialog.querySelector('[data-order-search]').addEventListener('click', searchOrders);
    query.addEventListener('keydown', event => {
        if (event.key !== 'Enter') return;
        event.preventDefault();
        void searchOrders();
    });
    dialog.addEventListener('close', () => { if (!dialog.open) pending?.abort(); });
}

const budgetLookup = document.querySelector('#commercial-budget-lookup');
if (budgetLookup) {
    const search = budgetLookup.querySelector('[data-sales-budget-search]');
    const query = search.querySelector('input');
    const status = budgetLookup.querySelector('[data-sales-budget-status]');
    const results = budgetLookup.querySelector('[data-sales-budget-results]');
    let pending;

    async function searchBudgets() {
        pending?.abort();
        const request = new AbortController();
        pending = request;
        results.replaceChildren();
        status.textContent = 'Buscando presupuestos…';
        try {
            const url = new URL(budgetLookup.dataset.budgetUrl, location.origin);
            url.searchParams.set('buscar', query.value.trim());
            const response = await fetch(url, { signal: request.signal });
            if (!response.ok || response.redirected) throw new Error();
            const budgets = await response.json();
            if (request.signal.aborted) return;
            for (const budget of budgets) {
                const row = results.insertRow();
                row.insertCell().textContent = budget.NumeroPresupuesto;
                row.insertCell().textContent = budget.Descripcion;
                const page = new URL(budgetLookup.dataset.budgetNewUrl, location.origin);
                page.searchParams.set('presupuestoId', budget.PresupuestoId);
                const select = document.createElement('a');
                select.href = page.href;
                select.className = 'hyt-button-secondary hyt-action-button';
                select.textContent = 'Cargar partidas';
                row.insertCell().append(select);
            }
            status.textContent = budgets.length === 100 ? 'Se muestran 100 presupuestos. Usa la búsqueda para acotar los resultados.'
                : budgets.length ? '' : 'No se encontraron presupuestos para este cliente.';
        } catch {
            if (!request.signal.aborted) status.textContent = 'No se pudo consultar la lista de presupuestos. Vuelve a intentarlo.';
        }
    }

    document.querySelector('[data-commercial-budget-open]').addEventListener('click', () => {
        setDialog(budgetLookup, true);
        void searchBudgets();
    });
    search.addEventListener('submit', event => {
        event.preventDefault();
        void searchBudgets();
    });
    budgetLookup.addEventListener('close', () => { if (!budgetLookup.open) pending?.abort(); });
}

const invoiceLookup = document.querySelector('#commercial-invoice-lookup');
if (invoiceLookup && form) {
    const search = invoiceLookup.querySelector('[data-credit-invoice-search]');
    const from = search.querySelector('#credit-invoice-from');
    const to = search.querySelector('#credit-invoice-to');
    const query = search.querySelector('#credit-invoice-query');
    const results = invoiceLookup.querySelector('[data-credit-invoice-results]');
    const status = invoiceLookup.querySelector('[data-credit-invoice-status]');
    const reference = form.querySelector('[data-credit-note-reference]');
    const format = new Intl.NumberFormat('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    let pending;

    async function searchInvoices() {
        pending?.abort();
        results.replaceChildren();
        if (!from.value || !to.value || from.valueAsNumber > to.valueAsNumber) {
            status.textContent = 'Indica un rango de fechas válido.';
            return;
        }
        const request = new AbortController();
        pending = request;
        status.textContent = 'Buscando facturas…';
        try {
            const url = new URL(invoiceLookup.dataset.invoiceUrl, location.origin);
            url.searchParams.set('desde', from.value);
            url.searchParams.set('hasta', to.value);
            url.searchParams.set('buscar', query.value.trim());
            const response = await fetch(url, { signal: request.signal });
            if (!response.ok || response.redirected) throw new Error();
            const invoices = await response.json();
            if (request.signal.aborted) return;
            for (const invoice of invoices) {
                const row = results.insertRow();
                row.insertCell().textContent = invoice.Referencia;
                row.insertCell().textContent = invoice.Fecha.split('-').reverse().join('/');
                row.insertCell().textContent = invoice.Moneda + ' ' + format.format(invoice.Importe);
                const select = document.createElement('button');
                select.type = 'button';
                select.className = 'hyt-button-secondary hyt-action-button';
                select.textContent = 'Seleccionar';
                select.addEventListener('click', () => {
                    reference.value = invoice.Referencia;
                    setDialog(invoiceLookup, false);
                });
                row.insertCell().append(select);
            }
            status.textContent = invoices.length === 20 ? 'Se muestran las 20 facturas más recientes. Ajusta el rango o busca por serie o número.'
                : invoices.length ? '' : 'No se encontraron facturas emitidas en este rango.';
        } catch {
            if (!request.signal.aborted) status.textContent = 'No se pudo consultar la lista de facturas. Vuelve a intentarlo.';
        }
    }

    form.querySelector('[data-credit-invoice-open]').addEventListener('click', () => {
        setDialog(invoiceLookup, true);
        void searchInvoices();
    });
    search.addEventListener('submit', event => {
        event.preventDefault();
        void searchInvoices();
    });
    invoiceLookup.addEventListener('close', () => { if (!invoiceLookup.open) pending?.abort(); });
}
