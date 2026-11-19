import { details } from './hyt-operational-map-details.js';

const tabs = [...document.querySelectorAll('[data-map-area]')];
const frame = document.querySelector('#hyt-map-frame');
const dialog = document.querySelector('#hyt-map-detail');
const links = document.querySelector('#hyt-map-module-links').content;
const root = document.documentElement;
const previousAdminNodes = { expediente: 'presupuesto', reportes: 'reportes-proveedores' };

let area;
let target;
let selectedNode;

function render() {
    const [requested, node] = location.hash.slice(1).split('/');
    const previousAdminNode = requested === 'administracion' && Object.hasOwn(previousAdminNodes, node);
    target = previousAdminNode ? previousAdminNodes[node] : node;
    let requestedArea = requested === 'planillas' || requested === 'obligaciones' ? 'planillas' : 'operacion';
    if (requestedArea === 'operacion' && !Object.hasOwn(details.operacion, target) && Object.hasOwn(details.planillas, target)) {
        requestedArea = 'planillas';
        if (target === 'telecredito') target = 'configuracion';
    }
    const tab = tabs.find(item => item.dataset.mapArea === requestedArea);
    area = tab.dataset.mapArea;
    const url = new URL(tab.dataset.mapUrl, location.href);
    url.searchParams.set('present', '1');
    url.searchParams.set('theme', root.dataset.hytTheme === 'dark' ? 'dark' : 'light');

    for (const item of tabs) {
        if (item === tab) item.setAttribute('aria-current', 'page');
        else item.removeAttribute('aria-current');
    }
    frame.title = `Mapa operativo de HYT · ${tab.textContent.trim()}`;
    if (frame.src !== url.href) frame.src = url.href;
    else focusDestination();
}

function focusDestination() {
    if (!target || frame.contentDocument?.readyState !== 'complete') return;
    const node = [...frame.contentDocument.querySelectorAll('[data-node-id]')]
        .find(item => item.dataset.nodeId === target);
    if (node) showDetail(node);
}

function showDetail(node) {
    const key = node.dataset.nodeId;
    const detail = details[area]?.[key];
    if (!detail) return;
    selectedNode = node;
    const { summary, rules, validation, result, modules } = detail;
    dialog.querySelector('h2').textContent = node.dataset.nodeLabel;
    dialog.querySelector('.hyt-map-detail-body').scrollTop = 0;
    dialog.querySelector('[data-map-summary]').textContent = summary;
    dialog.querySelector('[data-map-validation]').textContent = validation;
    dialog.querySelector('[data-map-result]').textContent = result;
    dialog.querySelector('.hyt-map-validation').open = false;
    dialog.querySelector('[data-map-rules]').replaceChildren(...rules.map(rule => {
        const item = document.createElement('li');
        item.textContent = rule;
        return item;
    }));
    const actions = dialog.querySelector('[data-map-actions]');
    actions.replaceChildren(...modules.flatMap(code => {
        const link = links.querySelector(`[data-module="${code}"]`);
        return link ? [link.cloneNode(true)] : [];
    }));
    if (!dialog.open) dialog.showModal();
}

frame.addEventListener('load', () => {
    const doc = frame.contentDocument;
    if (!doc) return;
    doc.addEventListener('click', event => {
        const node = event.target.closest?.('[data-node-id]');
        if (!node || !details[area]?.[node.dataset.nodeId]) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        showDetail(node);
    }, true);
    doc.addEventListener('keydown', event => {
        const node = event.target.closest?.('[data-node-id]');
        if (!node || !['Enter', ' '].includes(event.key)) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        showDetail(node);
    }, true);
    focusDestination();
});

dialog.addEventListener('close', () => selectedNode?.focus());
window.addEventListener('hashchange', render);
new MutationObserver(render).observe(root, { attributes: true, attributeFilter: ['data-hyt-theme'] });
render();
