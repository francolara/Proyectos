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
    for(const kind of ['guias','oc-cliente','ejecucion']){
      await navigate('/Presupuestos');
      const selector=kind==='oc-cliente'?'.hyt-action-registered[href^="/Seguimiento/oc-cliente/1/Editar"]':`.hyt-action-registered[href="/Seguimiento/${kind}/1"]`;
      if(!await evaluate(`document.querySelectorAll('.hyt-action-registered').length===3`))throw new Error('Deben resaltarse tres documentos registrados');
      if(!await evaluate(`getComputedStyle(document.querySelector(${JSON.stringify(selector)})).borderTopWidth==='3px'`))throw new Error('Falta borde registrado');
      await click(selector);
      if(kind==='oc-cliente'){
        await wait(`document.querySelector('#document-dialog [name="Fields[NumeroOCCliente]"]')`);
        if(!await evaluate(`document.querySelector('#document-dialog [name="Fields[NumeroOCCliente]"]').value==='OC-000A/01'&&document.querySelector('#document-dialog [name=returnUrl],#document-dialog [name=ReturnUrl]').value.includes('/Presupuestos')`))throw new Error('OC no usa edición y retorno de presupuestos');
        await evaluate(`const f=document.querySelector('#document-dialog form');f.querySelector('[name=MotivoCambio]').value='Prueba de edición desde presupuesto';f.requestSubmit()`);
        await wait(`location.pathname==='/Presupuestos'&&!document.querySelector('#document-dialog').open`);
        continue;
      }
      await wait(`location.pathname==='/Seguimiento/${kind}/1'`);
      if(!await evaluate(`document.body.textContent.includes('Datos del documento')||!!document.querySelector('[data-document-kind=GUIA] fieldset:disabled')`))throw new Error('Falta detalle de '+kind);
      if(kind!=='ejecucion'&&!await evaluate(`!!document.querySelector('a[href*="download=true"]')&&!!document.querySelector('a[target="_blank"][href*="/Archivos/"]')`))throw new Error('Falta acceso a adjuntos de '+kind);
    }
    await navigate('/Presupuestos');
    for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]){
      await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
      if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento '+label);
      const shot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview',`presupuesto-documentos-registrados-${label}.png`),Buffer.from(shot.data,'base64'));
    }
    console.log('OK: borde rojo en tres documentos, navegación a detalle y acceso a adjuntos, escritorio/móvil.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
