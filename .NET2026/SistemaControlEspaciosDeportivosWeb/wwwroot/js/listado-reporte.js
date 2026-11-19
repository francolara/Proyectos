// Firma: FRANCO LARA - 19/11/2026 | Imprime el listado y exporta filas, indicadores y totales visibles a CSV.
document.querySelector('[data-listado-imprimir]')?.addEventListener('click', () => window.print());

document.querySelector('[data-listado-csv]')?.addEventListener('click', () => {
    const report = document.querySelector('main');
    const listing = report.querySelector('[data-listado-tabla]');
    const textOf = element => element.innerText.trim().replace(/\s+/g, ' ');
    const encode = (text, isAmount = false) => {
        const numeric = isAmount && /^-?\d+(?:[.,]\d+)*$/.test(text);
        const value = !numeric && /^[=+\-@]/.test(text) ? "'" + text : text;
        return '"' + value.replaceAll('"', '""') + '"';
    };

    const tableLines = table => [...table.rows]
        .filter(row => !row.hasAttribute('data-listado-vacio'))
        .map(row => [...row.cells].flatMap(cell => [
            encode(textOf(cell), cell.matches('.listado-importe, .text-end')),
            ...Array(cell.colSpan - 1).fill('""')
        ]).join(';'));

    const lines = listing ? tableLines(listing) :
        [...report.querySelectorAll('h1, h2, .metric, .mini-metric, table')].flatMap(element => {
            if (element.matches('table')) return tableLines(element);
            const fields = element.matches('h1, h2') ? [element] : [...element.children];
            const values = fields.map(textOf).filter(Boolean);
            return values.length ? [values.map(value => encode(value)).join(';')] : [];
        });

    const file = new Blob(['\uFEFF', lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(file);
    const link = document.createElement('a');
    link.href = url;
    const title = report.querySelector('h1').textContent.trim();
    link.download = (listing ? listing.dataset.listadoTabla : title.replace(/[<>:"/\\|?*]/g, '-')) + '.csv';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});
