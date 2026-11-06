// Flujos MVC reales con repositorios de prueba aislados; nunca conecta SQL.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {setTimeout as delay} from 'node:timers/promises';
import {fileURLToPath} from 'node:url';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixture = JSON.parse(await fs.readFile(path.join(os.tmpdir(), 'hyt-maintenance-browser-fixture.json'), 'utf8'));
const targetUrl = fixture.url + '/Acceso';
if (!['127.0.0.1', 'localhost'].includes(new URL(targetUrl).hostname)) throw new Error('Solo se revisa el host local de pruebas.');
const tempRoot = await fs.realpath(os.tmpdir());
const profile = await fs.mkdtemp(path.join(tempRoot, 'hyt-browser-'));
const chrome = spawn('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'],
    {windowsHide: true, stdio: 'ignore'});
let socket;
try {
    let port;
    for (let i = 0; i < 100; i++) {
        try { port = (await fs.readFile(path.join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0]; break; }
        catch { await delay(100); }
    }
    if (!port) throw new Error('Chrome no publicó su canal de inspección.');
    const tab = await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, {method: 'PUT'})).json();
    socket = new WebSocket(tab.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {socket.addEventListener('open', resolve, {once: true}); socket.addEventListener('error', reject, {once: true});});
    let id = 0;
    const pending = new Map();
    const events = new Map();
    socket.addEventListener('message', message => {
        const result = JSON.parse(message.data);
        if (result.id) {
            const item = pending.get(result.id);
            if (item) { clearTimeout(item.timeout); pending.delete(result.id); result.error ? item.reject(new Error(JSON.stringify(result.error))) : item.resolve(result.result); }
        } else if (events.has(result.method)) { events.get(result.method)(result.params); events.delete(result.method); }
    });
    const send = (method, params = {}) => new Promise((resolve, reject) => {
        const requestId = ++id;
        const timeout = setTimeout(() => { pending.delete(requestId); reject(new Error(`Tiempo agotado: ${method}`)); }, 15000);
        pending.set(requestId, {resolve, reject, timeout}); socket.send(JSON.stringify({id: requestId, method, params}));
    });
    await send('Page.enable');

    async function evaluate(expression) {
        const result = await send('Runtime.evaluate', {expression, awaitPromise:true, returnByValue:true});
        if(result.exceptionDetails) throw new Error(result.exceptionDetails.text + ' ' + result.exceptionDetails.exception?.description);
        return result.result.value;
    }
    async function wait(expression) {
        for(let i=0;i<100;i++){ if(await evaluate(expression)) return; await delay(100); }
        throw new Error('No se cumplió: '+expression+' Estado: '+await evaluate("JSON.stringify({url:location.href,dialog:document.querySelector(\"#document-dialog\")?.textContent,open:document.querySelector(\"#document-dialog\")?.open,validation:document.querySelector('.validation-summary-errors')?.textContent,invalid:[...document.querySelectorAll('[data-budget-form] input,[data-budget-form] select,[data-budget-form] textarea')].filter(e=>!e.checkValidity()).map(e=>({name:e.name,value:e.value,message:e.validationMessage}))})"));
    }
    async function navigate(route) {
        const loaded = new Promise(resolve=>events.set('Page.loadEventFired',resolve));
        await send('Page.navigate',{url:fixture.url+route}); await loaded;
        await wait(`location.pathname===${JSON.stringify(route.split('?')[0])} && document.readyState==='complete'`);
    }
    async function click(selector) {
        const point=await evaluate(`(()=>{const element=document.querySelector(${JSON.stringify(selector)});element.scrollIntoView({block:'center'});const rect=element.getBoundingClientRect();return {x:rect.x+rect.width/2,y:rect.y+rect.height/2};})()`);
        await send('Input.dispatchMouseEvent',{type:'mousePressed',button:'left',clickCount:1,...point});
        await send('Input.dispatchMouseEvent',{type:'mouseReleased',button:'left',clickCount:1,...point});
    }


    async function frameClick(selector){const point=await evaluate(`(()=>{const frame=document.querySelector('[data-material-price-frame]'),element=frame.contentDocument.querySelector(${JSON.stringify(selector)});element.scrollIntoView({block:'center'});const r=element.getBoundingClientRect(),f=frame.getBoundingClientRect();return {x:r.x+f.x+r.width/2,y:r.y+f.y+r.height/2};})()`);await send('Input.dispatchMouseEvent',{type:'mousePressed',button:'left',clickCount:1,...point});await send('Input.dispatchMouseEvent',{type:'mouseReleased',button:'left',clickCount:1,...point});}
    const frameDoc="document.querySelector('[data-material-price-frame]').contentDocument";
    async function frameReady(field){await wait(`document.querySelector('#material-price-dialog').open && ${frameDoc}?.readyState==='complete' && ${frameDoc}.querySelector(${JSON.stringify(field)})`);}
    async function capture(name){const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});await fs.writeFile(path.join(project,'entregables','preview',name+'.png'),Buffer.from(screenshot.data,'base64'));}
    await navigate('/Acceso');await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('form').requestSubmit()`);await wait("location.pathname==='/'");
    await send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
    await navigate('/Presupuestos/Nuevo');
    if(!await evaluate("document.querySelector('.hyt-workspace-wide') && !document.querySelector('[data-tariff-open]') && !document.querySelector('[data-apply-budget-defaults]') && document.querySelector('[data-manual-line]').textContent==='Agregar ítem' && document.querySelector('[data-manual-line]').classList.contains('hyt-button-primary')"))throw new Error('No se simplificó la cabecera de detalle.');
    if(!await evaluate("document.querySelector('#budget-record-number').textContent==='Se asigna al guardar' && getComputedStyle(document.querySelector('.hyt-internal-code')).borderTopWidth==='4px'"))throw new Error('Falta tarjeta de número de presupuesto.');
    await click('[data-manual-line]');await click('[data-line-material]');await wait("document.querySelector('[data-budget-lookup-body] button')");await click('[data-budget-lookup-body] button');await wait("document.querySelector('[name$=\".Unidad\"]').value==='UND'");
    await click('[data-replace-tariff]');await wait("document.querySelector('#budget-lookup').open && document.querySelector('[data-budget-lookup-body] button')");
    if(!await evaluate("document.querySelector('[name=\"Draft.ClienteId\"]').value===''"))throw new Error('Prueba debe abrir tarifa sin cliente.');
    await click('[data-budget-lookup-body] button');
    if(!await evaluate("[...document.querySelectorAll('[data-line-material],[data-replace-tariff],[data-line-service],[data-remove-line]')].every(b=>b.querySelector('svg') && b.getAttribute('aria-label') && b.getBoundingClientRect().width<=36) && document.querySelector('[name$=\".NumeroNecesidad\"]').getBoundingClientRect().width<=48"))throw new Error('Controles del detalle no son compactos.');
    await click('[data-remove-line]');
    await click('[data-budget-reference=ClienteId] [data-reference-open]');await wait("document.querySelector('[data-budget-lookup-body] button')");await click('[data-budget-lookup-body] button');
    await click('[data-budget-reference=TiendaId] [data-reference-open]');await wait("document.querySelector('[data-budget-lookup-body] button')");await click('[data-budget-lookup-body] button');
    await evaluate("document.querySelector('[name=\"Draft.Descripcion\"]').value='Detalle de servicio de prueba'");
    await click('[data-manual-line]');
    if(!await evaluate("document.querySelector('[name$=\".Cantidad\"]').value==='' && !document.querySelector('[data-origin-price]').readOnly"))throw new Error('Cantidad no inicia vacía o precio bloqueado.');
    await click('[data-line-material]');await wait("document.querySelector('[data-budget-lookup-body] button')");await click('[data-budget-lookup-body] button');
    await wait("document.querySelector('[name$=\".Unidad\"]').value==='UND'");
    await click('[data-replace-tariff]');await wait("document.querySelector('[data-budget-lookup-body] button')");
    if(!await evaluate("document.querySelector('#budget-lookup-title').textContent==='Seleccionar tarifa del material' && document.querySelector('[data-budget-lookup-body]').textContent.includes('100.00')"))throw new Error('Tarifa con IGV no se muestra neta.');
    await click('[data-budget-lookup-body] button');
    await evaluate("document.querySelector('[name$=\".Cantidad\"]').value='2.50';document.querySelector('[data-origin-price]').value='80.00';document.querySelector('[name$=\".NumeroNecesidad\"]').value='NEC-A1';document.querySelector('[data-origin-price]').dispatchEvent(new Event('input',{bubbles:true}))");
    await click('[data-line-service]');await wait("document.querySelector('[data-budget-lookup-body] button')");await click('[data-budget-lookup-body] button');
    await wait("document.querySelector('[data-line-total]').textContent==='200.00' && document.querySelector('[data-budget-tax]').textContent==='36.00' && document.querySelector('[data-budget-total]').textContent==='236.00'");
    await capture('presupuesto-detalle-compacto-escritorio');
    await evaluate("document.querySelector('[name=\"Draft.Revision[CodigoMoneda]\"]').value='USD';document.querySelector('[name=\"Draft.Revision[CodigoMoneda]\"]').dispatchEvent(new Event('change',{bubbles:true}))");
    await wait("document.querySelector('[data-origin-price]').value===''");
    await evaluate("document.querySelector('[data-origin-price]').value='10.00';document.querySelector('[data-origin-price]').dispatchEvent(new Event('input',{bubbles:true}))");
    await wait("document.querySelector('[data-line-total]').textContent==='25.00'");
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
    if(!await evaluate('document.documentElement.scrollWidth<=innerWidth+1'))throw new Error('Detalle desborda fuera de su tabla en móvil.');
    await capture('presupuesto-detalle-compacto-movil');
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate('/Maestros/clientes');
    if(!await evaluate("(()=>{const buttons=[...document.querySelectorAll('.hyt-record-action')];return buttons.length>=2 && buttons.every(b=>Math.abs(b.getBoundingClientRect().width-84)<1);})()"))throw new Error('Ver y Editar tienen distintos anchos.');
    await navigate('/Maestros/clientes/Nuevo');
    if(!await evaluate("getComputedStyle(document.querySelector('.hyt-internal-code')).borderTopWidth==='4px'"))throw new Error('Código del cliente sin tarjeta.');
    console.log('OK navegador: iconos compactos, tarifa sin cliente, botones uniformes, tarjetas y ancho completo, agregar ítem, material/unidad, tarifa neta, cantidad explícita, precio editable, necesidad/servicio, subtotal/IGV y moneda sin precio; móvil sin desborde.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
