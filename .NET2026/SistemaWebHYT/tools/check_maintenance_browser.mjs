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
        throw new Error('No se cumplió: '+expression+' Estado: '+await evaluate("JSON.stringify({url:location.href,status:document.querySelector('[data-lookup-status]')?.textContent,dialog:document.querySelector('#master-lookup')?.outerHTML.slice(-600)})"));
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
    await navigate('/Maestros/materiales/Nuevo');
    await click('[data-open-zones]');
    await wait("document.querySelector('#master-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-lookup-rows] input')");
    await click('[data-lookup-rows] input');
    await click('[data-lookup-next]');
    await wait("document.querySelector('[data-lookup-page]').textContent.includes('Página 2') && document.querySelector('#master-lookup').getAttribute('aria-busy')==='false'");
    await click('[data-lookup-rows] input');
    await click('[data-lookup-prev]');
    await wait("document.querySelector('[data-lookup-page]').textContent.includes('Página 1') && document.querySelector('#master-lookup').getAttribute('aria-busy')==='false'");
    if(!await evaluate("document.querySelector('[data-lookup-rows] input').checked && document.querySelectorAll('[data-zone-values] input:checked').length===2")) throw new Error('Selección múltiple perdida entre páginas.');
    await click('[data-close-lookup]');
    await click('[data-reference="UnidadMedidaId"] [data-open-lookup]');
    await wait("document.querySelector('#master-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-lookup-rows] button')");
    await click('[data-lookup-rows] button');
    await wait("!document.querySelector('#master-lookup').open");
    if(!await evaluate("document.querySelector('[name=\"Fields[UnidadMedidaId]\"]').value==='1' && document.activeElement.matches('[data-reference=\"UnidadMedidaId\"] [data-open-lookup]')")) throw new Error('Selección o retorno de foco incorrecto.');
    console.log('OK navegador: zonas persistentes entre páginas y ayuda de unidad con retorno de foco.');
    for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]) {
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});
        await evaluate('new Promise(requestAnimationFrame)');
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1')) throw new Error('Desbordamiento en ficha '+label);
        const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
        const output=path.join(project,'entregables','preview',`material-${label}.png`);
        await fs.mkdir(path.dirname(output),{recursive:true}); await fs.writeFile(output,Buffer.from(shot.data,'base64'));
    }
    await navigate('/Maestros/partidas-tarifario/Nuevo?parentId=1');
    if(!await evaluate("document.querySelector('[name=\"Fields[CodigoMoneda]\"]').value==='PEN' && document.querySelector('[name=\"Fields[CodigoTipoIGV]\"]').value==='SIN_IGV' && document.querySelectorAll('[data-zone-values] input:checked').length===2")) throw new Error('No se cargaron predeterminados del padre.');
    await click('[data-reference="MaterialId"] [data-open-lookup]');
    await wait("document.querySelector('#master-lookup').getAttribute('aria-busy')==='false' && document.querySelector('[data-lookup-rows] button')");
    await click('[data-lookup-rows] button');
    await evaluate("document.querySelector('[name=\"Fields[Posicion]\"]').value='POS-TARIFA'");
    await click('[data-material-codes]');
    await wait("document.querySelector('[data-defaults-status]').textContent.includes('Códigos copiados')");
    if(!await evaluate("document.querySelector('[name=\"Fields[Posicion]\"]').value==='POS-TARIFA' && document.querySelector('[name=\"Fields[CodigoMaterialSAP]\"]').value==='000SAP' && document.querySelector('[name=\"Fields[PrecioUSD]\"]').value===''") ) throw new Error('Códigos o USD incorrectos.');
    console.log('OK navegador: predeterminados por versión, códigos alfanuméricos y USD vacío.');
    await navigate('/Configuracion/TasasIGV');
    if(!await evaluate("document.body.textContent.includes('0.18') && !!document.querySelector('.app-date-picker-trigger')")) throw new Error('Falta tasa o calendario.');
    if(await evaluate('document.documentElement.scrollWidth>innerWidth+1')) throw new Error('Desbordamiento en tasas móvil.');
    await navigate('/Maestros/clientes?search=abc&page=2');
    if(await evaluate('document.documentElement.scrollWidth>innerWidth+1')) throw new Error('Desbordamiento listado móvil.');
    const viewHref=await evaluate("document.querySelector('table a').getAttribute('href')");
    if(!viewHref.includes('returnUrl=')) throw new Error('Listado pierde retorno.');
    console.log('OK navegador: tasas IGV, calendario y listado móvil sin desbordamiento.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
