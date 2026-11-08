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
        for(let i=0;i<100;i++){ if(await evaluate("!!document.body && ("+expression+")")) return; await delay(100); }
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

    await navigate('/Acceso');
    await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('form').requestSubmit()`);
    await wait("location.pathname==='/'");
    await navigate('/Seguimiento/guias');
    if(!await evaluate(`![...document.querySelectorAll('.hyt-table th')].some(e=>e.textContent==='Estado')&&!document.querySelector('select[name=state]')`))throw new Error('Listado guía conserva estados');
    await click('a[href="/Seguimiento/guias/1"]');await wait(`location.pathname==='/Seguimiento/guias/1'`);
    const values=await evaluate(`(()=>{const f=document.querySelector('[data-document-kind=GUIA]');return [...f.querySelectorAll('[name^="Fields["]')].map(e=>[e.name,e.value])})()`);
    if(!await evaluate(`!!document.querySelector('fieldset:disabled')&&!document.querySelector('#document-state')&&!document.querySelector('button[type=submit]')&&!document.querySelector('input[type=file]')&&!document.querySelector('[name=RemovedFiles]')&&!!document.querySelector('a[href*="download=true"]')`))throw new Error('Consulta guía permite edición o muestra estados');
    const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});await fs.writeFile(path.join(project,'entregables','preview','guia-consulta-sin-estados.png'),Buffer.from(shot.data,'base64'));
    await navigate('/Seguimiento/guias');await click('[data-document-open][href*="/Editar"]');await wait(`document.querySelector('#document-dialog [data-document-kind=GUIA]')`);
    const editing=await evaluate(`(()=>{const f=document.querySelector('#document-dialog [data-document-kind=GUIA]');return [...f.querySelectorAll('[name^="Fields["]')].map(e=>[e.name,e.value])})()`);
    if(JSON.stringify(values)!==JSON.stringify(editing))throw new Error('Consulta y edición no comparten campos/valores');
    if(!await evaluate(`!document.querySelector('#document-dialog fieldset').disabled&&!!document.querySelector('#document-dialog button[type=submit]')&&!!document.querySelector('#document-dialog input[type=file]')&&!document.querySelector('#document-state')`))throw new Error('Edición guía incorrecta');
    await click('#document-dialog [data-hyt-close]');
    await navigate('/Seguimiento/oc-cliente');await click('.hyt-cell-action[href="/Seguimiento/guias/1/Adjuntos"]');await wait(`document.querySelector('#document-dialog fieldset:disabled')`);
    for(const [width,height] of [[1440,1000],[390,844]]){await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento consulta guía');}
    console.log('OK: Guía sin estados; Ver/Editar comparten datos y campos, consulta bloqueada y adjuntos disponibles, escritorio/móvil.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
