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
        throw new Error('No se cumplió: '+expression+' Estado: '+await evaluate("JSON.stringify({url:location.href,validation:document.querySelector('.validation-summary-errors')?.textContent,invalid:[...document.querySelectorAll('[data-budget-form] input,[data-budget-form] select,[data-budget-form] textarea')].filter(e=>!e.checkValidity()).map(e=>({name:e.name,value:e.value,message:e.validationMessage}))})"));
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

    await navigate('/Acceso');
    await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('form').requestSubmit()`);
    await wait("location.pathname==='/'");
    await navigate('/Presupuestos/Nuevo');
    await click('[data-budget-reference="ClienteId"] [data-reference-open]');
    await wait("document.querySelector('#budget-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-budget-lookup-body] button')");
    await click('[data-budget-lookup-body] button');
    await wait("!document.querySelector('#budget-lookup').open");
    await click('[data-budget-reference="TiendaId"] [data-reference-open]');
    await wait("document.querySelector('#budget-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-budget-lookup-body] button')");
    await click('[data-budget-lookup-body] button');
    await wait("!document.querySelector('#budget-lookup').open");
    await click('[data-tariff-open]');
    await wait("document.querySelector('#budget-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-budget-lookup-body] button')");
    if(!await evaluate("document.querySelector('[data-budget-lookup-body]').textContent.includes('Sur, Centro')"))throw new Error('Ayuda no muestra zonas.');
    await click('[data-budget-lookup-body] button');
    await wait("document.querySelectorAll('[data-budget-line]').length===1 && document.querySelector('[data-budget-total]').textContent==='118.00'");
    await evaluate(`document.querySelector('[name="Draft.Partidas[0].Cantidad"]').value='2.00';document.querySelector('[name="Draft.Partidas[0].Cantidad"]').dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#budget-description').value='Servicio de prueba';`);
    await wait("document.querySelector('[data-budget-total]').textContent==='236.00'");
    await evaluate(`document.querySelector('[name="Draft.Revision[CodigoMoneda]"]').value='USD';document.querySelector('[name="Draft.Revision[CodigoMoneda]"]').dispatchEvent(new Event('change',{bubbles:true}));`);
    await wait("document.querySelector('[data-budget-total]').textContent==='Pendiente'");
    if(!await evaluate("document.querySelector('[name=\"Draft.Partidas[0].PrecioManual\"]').value===''"))throw new Error('Cambio de moneda dejó precio anterior.');
    await click('[data-replace-tariff]');
    await wait("document.querySelector('#budget-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-budget-lookup-body] button')");
    if(!await evaluate("document.querySelector('[data-budget-lookup-body]').textContent.includes('Precio pendiente')"))throw new Error('USD ausente no se identifica.');
    await click('[data-budget-lookup-body] button');
    await evaluate(`const origin=document.querySelector('[name="Draft.Partidas[0].OrigenPrecio"]');origin.value='MANUAL';origin.dispatchEvent(new Event('change',{bubbles:true}));document.querySelector('[name="Draft.Partidas[0].PrecioManual"]').value='50.00';document.querySelector('[name="Draft.Partidas[0].MotivoPrecioManual"]').value='Precio acordado en dólares';document.querySelector('[name="Draft.Partidas[0].PrecioManual"]').dispatchEvent(new Event('input',{bubbles:true}));`);
    await wait("document.querySelector('[data-budget-total]').textContent==='118.00'");
    console.log('OK navegador presupuesto: cliente/tienda, ayuda con zonas, cálculo y USD ausente sin conversión.');
    for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]){
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');await evaluate("document.activeElement?.blur();document.querySelector('.hyt-budget-lines .hyt-table-wrap').scrollLeft=0");
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento presupuesto '+label+' '+await evaluate("JSON.stringify({width:innerWidth,doc:document.documentElement.scrollWidth,body:document.body.scrollWidth,x:scrollX,items:[...document.querySelectorAll('body *')].filter(e=>e.getBoundingClientRect().right>innerWidth+1&&!e.parentElement?.closest('.hyt-table-wrap')).slice(0,14).map(e=>({tag:e.tagName,cls:e.className,w:e.getBoundingClientRect().width,right:e.getBoundingClientRect().right,scroll:e.scrollWidth}))})"));
        const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});const output=path.join(project,'entregables','preview',`presupuesto-${label}.png`);await fs.mkdir(path.dirname(output),{recursive:true});await fs.writeFile(output,Buffer.from(screenshot.data,'base64'));
    }
    await evaluate("document.querySelector('[data-budget-form]').requestSubmit()");
    await wait("location.pathname==='/Presupuestos/3' && document.readyState==='complete' && !!document.querySelector('fieldset:disabled')");
    if(!await evaluate("!!document.querySelector('fieldset:disabled') && !document.body.textContent.includes('Guardar borrador')"))throw new Error('Consulta permite editar después de guardar.');
    console.log('OK navegador presupuesto: escritorio/móvil sin desbordamiento y guardado MVC con consulta posterior.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
