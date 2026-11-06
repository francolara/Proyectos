// Firma: FRANCO LARA - 03/11/2026 | Ayudas MVC, zonas múltiples y predeterminados por versión/material.
// Firma: FRANCO LARA - 05/11/2026 | Copia SAP, unidades por abreviatura y campos por persona, número de documento y tipo de identidad oculto y títulos específicos en ayudas.
import { setDialog } from './hyt-ui.js';

const dialog = document.querySelector('#master-lookup');
if (dialog) {
    const searchForm = dialog.querySelector('[data-lookup-search]');
    const rows = dialog.querySelector('[data-lookup-rows]');
    const status = dialog.querySelector('[data-lookup-status]');
    const previous = dialog.querySelector('[data-lookup-prev]');
    const next = dialog.querySelector('[data-lookup-next]');
    const zoneValues = document.querySelector('[data-zone-values]');
    const selected = new Map();
    zoneValues?.querySelectorAll('input:checked').forEach(input => selected.set(Number(input.value), input.closest('label').textContent.trim()));
    let target = null, entity = '', page = 1, busy = false, request = 0;
    function renderZones() {
        if (!zoneValues) return;
        zoneValues.replaceChildren();
        for (const [id, label] of selected) {
            const wrapper = document.createElement('label');
            wrapper.dataset.zoneItem = String(id);
            const input = document.createElement('input');
            input.type = 'checkbox'; input.name = 'ZoneIds'; input.value = String(id); input.checked = true;
            input.addEventListener('change', () => { selected.delete(id); renderZones(); });
            wrapper.append(input, document.createTextNode(label)); zoneValues.append(wrapper);
        }
    }
    function close() { if (!busy) { request++; setDialog(dialog, false); } }
    async function applyDefaults(material = false) {
        const feedback = document.querySelector('[data-defaults-status]');
        const field = material ? 'MaterialId' : 'VersionId';
        const id = document.querySelector(`[data-reference="${field}"] input`)?.value;
        if (!feedback) return;
        if (!id) { feedback.textContent = material ? 'Selecciona un material primero.' : 'Selecciona una versión primero.'; return; }
        const button = document.querySelector(material ? '[data-material-codes]' : '[data-apply-defaults]');
        if (button) button.disabled = true;
        feedback.textContent = 'Consultando valores…';
        try {
            const source = material ? 'materiales' : 'versiones-tarifario';
            const response = await fetch(`/Maestros/${source}/${encodeURIComponent(id)}/Predeterminados`, {headers:{Accept:'application/json'}});
            if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error();
            const data = await response.json();
            if (document.querySelector(`[data-reference="${field}"] input`)?.value !== id) return;
            const names = material ? {CodigoMaterialSAP:'CodigoMaterialSAP'} : {CodigoMonedaPredeterminada:'CodigoMoneda',CodigoTipoIGVPredeterminado:'CodigoTipoIGV'};
            for (const [sourceName, destination] of Object.entries(names)) {
                const input = document.querySelector(`[name="Fields[${destination}]"]`);
                if (input) input.value = data.fields[sourceName] ?? '';
            }
            if (!material) {
                selected.clear();
                for (const zone of data.zones ?? []) selected.set(zone.Id, zone.Descripcion ?? `Zona ${zone.Id}`);
                renderZones();
            }
            feedback.textContent = material ? 'Códigos copiados. Puedes ajustarlos para esta partida.' : 'Moneda, tratamiento IGV y zonas aplicados. Revisa los precios de la partida.';
        } catch { feedback.textContent = 'No se pudieron consultar los valores. Reintenta la operación.'; }
        finally { if (button) button.disabled = false; }
    }
    async function load() {
        const current = ++request;
        busy = true; rows.replaceChildren(); status.textContent = 'Cargando registros…';
        previous.disabled = true; next.disabled = true;
        setDialog(dialog, true, true, close);
        try {
            const parameters = new URLSearchParams({ search: searchForm.elements.search.value, page: String(page) });
            // Las tiendas ofrecidas a un contacto pertenecen al cliente elegido.
            if (entity === 'tiendas') {
                const client = document.querySelector('[data-reference="ClienteId"] input')?.value;
                if (client) parameters.set('parentId', client);
            }
            const response = await fetch(`/Maestros/${encodeURIComponent(entity)}/Ayuda?${parameters}`, { headers: { Accept: 'application/json' } });
            if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error();
            const data = await response.json();
            if (current !== request) return;
            for (const row of data.rows) {
                const tr = document.createElement('tr');
                for (const text of [row.code, row.description, row.extra ?? '']) {
                    const cell = document.createElement('td'); cell.textContent = text; tr.append(cell);
                }
                const cell = document.createElement('td');
                if (target === null) {
                    const input = document.createElement('input'); input.type = 'checkbox'; input.checked = selected.has(row.id);
                    input.setAttribute('aria-label', `Seleccionar ${row.description}`);
                    input.addEventListener('change', () => {
                        if (input.checked) selected.set(row.id, `${row.code} · ${row.description}`); else selected.delete(row.id);
                        renderZones();
                    });
                    cell.append(input);
                } else {
                    const button = document.createElement('button'); button.type = 'button'; button.className = 'hyt-button-secondary'; button.textContent = 'Seleccionar';
                    button.addEventListener('click', () => {
                        target.querySelector('input').value = String(row.id);
                        target.querySelector('[data-reference-label]').textContent = `${row.code} · ${row.description}`;
                        if (target.dataset.reference === 'VersionId' && document.querySelector('[data-apply-defaults]')) applyDefaults();
                        if (target.dataset.reference === 'ClienteId') {
                            const store = document.querySelector('[data-reference="TiendaId"]');
                            if (store) { store.querySelector('input').value = ''; store.querySelector('[data-reference-label]').textContent = 'Sin seleccionar'; }
                        }
                        close();
                    });
                    cell.append(button);
                }
                tr.append(cell); rows.append(tr);
            }
            status.textContent = data.rows.length ? `${data.total} registros encontrados.` : 'No se encontraron registros.';
            dialog.querySelector('[data-lookup-page]').textContent = `Página ${page} · 20 por página`;
            previous.disabled = page <= 1; next.disabled = page * 20 >= data.total;
        } catch {
            if (current === request) status.textContent = 'No se pudo cargar la ayuda. Vuelve a buscar para reintentar.';
        } finally {
            if (current === request) { busy = false; setDialog(dialog, true, false, close); }
        }
    }
    function open(reference, key) {
        const heading=dialog.querySelector('[data-lookup-code-heading]');if(heading)heading.textContent=key==='unidades'?'Abreviatura':key==='versiones-tarifario'?'Versión':'Código interno';
        target = reference; entity = key; page = 1; searchForm.reset();
        dialog.querySelector('#lookup-title').textContent = reference ? ({unidades:'Seleccionar unidad de medida',clientes:'Seleccionar cliente',tiendas:'Seleccionar tienda',materiales:'Seleccionar material o servicio',especialidades:'Seleccionar especialidad','tipos-servicio':'Seleccionar tipo de servicio','versiones-tarifario':'Seleccionar versión de tarifario',tarifarios:'Seleccionar tarifario'}[key] ?? 'Seleccionar registro') : 'Seleccionar zonas';
        load();
    }
    document.querySelectorAll('[data-open-lookup]').forEach(button => button.addEventListener('click', () => {
        const reference = button.closest('[data-reference]'); open(reference, reference.dataset.lookup);
    }));
    document.querySelectorAll('[data-clear-reference]').forEach(button => button.addEventListener('click', () => {
        const reference = button.closest('[data-reference]'); reference.querySelector('input').value = '';
        reference.querySelector('[data-reference-label]').textContent = 'Sin seleccionar';
    }));
    document.querySelector('[data-open-zones]')?.addEventListener('click', () => open(null, 'zonas'));
    document.querySelector('[data-apply-defaults]')?.addEventListener('click', () => applyDefaults());
    document.querySelector('[data-material-codes]')?.addEventListener('click', () => applyDefaults(true));
    dialog.querySelectorAll('[data-close-lookup]').forEach(button => button.addEventListener('click', close));
    searchForm.addEventListener('submit', event => { event.preventDefault(); if (!busy) { page = 1; load(); } });
    previous.addEventListener('click', () => { if (!busy && page > 1) { page--; load(); } });
    next.addEventListener('click', () => { if (!busy) { page++; load(); } });
}

const clientForm = document.querySelector('[data-client-form="true"]');
if (clientForm) {
    const type = clientForm.querySelector('[name="Fields[TipoEntidad]"]');
    const number = clientForm.querySelector('[name="Fields[NumeroIdentificacion]"]');
    const readOnly = clientForm.dataset.readonly === 'true';
    function updateClientIdentity() {
        const legal = type.value === 'JURIDICA', natural = type.value === 'NATURAL';
        for (const name of ['RazonSocial', 'Nombres', 'Apellidos']) {
            const wrapper = clientForm.querySelector(`[data-field="${name}"]`);
            const input = wrapper.querySelector('input');
            const enabled = name === 'RazonSocial' ? legal : natural;
            input.disabled = !enabled;
            input.required = enabled && !readOnly;
            wrapper.querySelector('label').textContent = ({RazonSocial:'Razón social', Nombres:'Nombres', Apellidos:'Apellidos'})[name] + (enabled && !readOnly ? ' *' : '');
        }
        number.readOnly = readOnly || (!legal && !natural);
        number.required = !readOnly && (legal || natural);
        const label = clientForm.querySelector('[data-field="NumeroIdentificacion"] label');
        label.textContent = (legal ? 'Número de RUC' : 'Número de DNI / CE') + (number.required ? ' *' : '');
    }
    type.addEventListener('change', updateClientIdentity);
    updateClientIdentity();
}
