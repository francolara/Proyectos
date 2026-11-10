// Firma: FRANCO LARA - 10/11/2026 | Presenta el estimado mensual y su detalle de trabajadores en la ficha de obligación.
const estimates = [...document.querySelectorAll('[data-obligation-workers]')];
const workersMessage = document.querySelector('[data-obligation-workers-message]');
if (workersMessage) {
    const type = document.querySelector('#payrollobligations-type-1');
    const entity = document.querySelector('#payrollobligations-entity-1');
    const amount = document.querySelector('#payrollobligations-amount-1');
    const status = document.querySelector('[data-obligation-estimate]');
    const apply = document.querySelector('[data-obligation-use-estimate]');
    const money = new Intl.NumberFormat('es-PE', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    let selected;

    function refresh() {
        const name = entity.value.trim().toLocaleLowerCase('es-PE');
        const group = estimates.find(row => row.dataset.regime === type.value &&
            row.dataset.entity.toLocaleLowerCase('es-PE') === name);
        selected = group?.dataset;
        for (const row of estimates) row.hidden = row !== group;
        workersMessage.hidden = Boolean(group);
        if (!status) return;
        status.textContent = selected
            ? `Base mensual S/ ${money.format(selected.base)} · ${selected.rate} % · Estimado S/ ${money.format(selected.amount)}`
            : 'Selecciona una entidad del catálogo para ver el estimado.';
        apply.hidden = !selected || Number(selected.amount) <= 0;
    }

    entity.addEventListener('input', refresh);
    entity.addEventListener('change', refresh);
    type.addEventListener('change', refresh);
    apply?.addEventListener('click', () => {
        amount.value = selected.amount;
        amount.focus();
    });
    refresh();
}
