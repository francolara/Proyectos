// Firma: FRANCO LARA - 10/11/2026 | Impresión y CSV para Excel con títulos, detalle y totales en celdas separadas.
document.querySelector('[data-report-print]')?.addEventListener('click', () => window.print());

document.querySelector('[data-report-csv]')?.addEventListener('click', () => {
    const encodeCell = value => {
        const numeric = /^[-+]\d+(?:[.,]\d+)*$/.test(value);
        const safe = !numeric && /^[=+\-@]/.test(value) ? "'" + value : value;
        return '"' + safe.replaceAll('"', '""') + '"';
    };

    const separator = ';';
    const content = document.querySelectorAll(
        'main table, main h1, main h2, main h3, main h4, main p, main .new-brand strong, main .new-signature strong, main footer'
    );
    const sections = [...content].flatMap(element => {
        if (element.matches('table')) {
            const rows = [...element.rows]
                .filter(row => !row.matches('.old-empty-row, .new-empty-row'))
                .map(row => [...row.cells].flatMap(cell =>
                    [encodeCell(cell.innerText.trim()), ...Array(cell.colSpan - 1).fill('""')]))
                .filter(row => row.some(cell => cell !== '""'));
            return rows.length ? [rows.map(row => row.join(separator)).join('\r\n') + '\r\n'] : [];
        }
        if (element.closest('table') || element.querySelector('p, h1, h2, h3, h4')) return [];
        const value = element.innerText.trim();
        return value ? [encodeCell(value)] : [];
    });
    const csv = new Blob(['\uFEFF', 'sep=', separator, '\r\n', sections.join('\r\n')], {
        type: 'text/csv;charset=utf-8'
    });
    const url = URL.createObjectURL(csv);
    const link = document.createElement('a');
    link.href = url;
    link.download = (document.title.trim().replace(/[<>:"/\\|?*]/g, '-') || 'Reporte HYT') + '.csv';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});
