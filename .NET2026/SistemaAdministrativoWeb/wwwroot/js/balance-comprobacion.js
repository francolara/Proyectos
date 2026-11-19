// Firma: FRANCO LARA - 18/11/2026 | Exporta tablas, agrupaciones y totales de los reportes HTML a CSV.
document.querySelector('[data-balance-print]')?.addEventListener('click', () => window.print());

document.querySelector('[data-balance-csv]')?.addEventListener('click', () => {
    const report = document.querySelector('.balance-document');
    const escapeCell = (value, isAmount = false) => {
        const numericAmount = isAmount && /^[+-]?[\d.,\s]+$/.test(value);
        const safeValue = !numericAmount && /^[=+@-]/.test(value) ? "'" + value : value;
        return '"' + safeValue.replaceAll('"', '""') + '"';
    };

    const tableRows = table => {
        let headers = [...table.tHead.querySelectorAll('[scope="col"]')];
        if (headers.length === 0) {
            headers = [...table.tHead.rows[0].cells];
        }
        const rows = [...table.rows]
            .filter(row => row.parentElement !== table.tHead && !row.hasAttribute('data-report-csv-skip'));

        return [
            headers.map(cell => escapeCell(cell.dataset.csvLabel ?? cell.textContent.trim())),
            ...rows.map(row => [...row.cells].flatMap(cell => [
                escapeCell(cell.textContent.trim(), cell.classList.contains('balance-amount')),
                ...Array(cell.colSpan - 1).fill('""')
            ]))
        ].map(row => row.join(';'));
    };

    const rows = [...report.querySelectorAll('.balance-document-table, [data-report-csv-context]')]
        .flatMap(element => {
            if (element.matches('table')) {
                return tableRows(element);
            }
            const context = [...element.children].map(child =>
                child.innerText.trim().replace(/\s*\n\s*/g, ' ').replace(/\s{2,}/g, ' '));
            return context.some(Boolean) ? [context.map(value => escapeCell(value)).join(';')] : [];
        });

    const csv = new Blob(['\uFEFF', rows.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(csv);
    const link = document.createElement('a');
    link.href = url;
    const title = report.querySelector('#balance-document-title').textContent.trim();
    link.download = (title === 'Balance de comprobación mensual' ? 'Balance de comprobación' : title) + '.csv';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});
