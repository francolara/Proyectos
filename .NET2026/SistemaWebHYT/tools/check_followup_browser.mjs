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
    for(const route of ['/Seguimiento/oc-cliente','/Seguimiento/guias','/Seguimiento/oc-cliente/1','/Seguimiento/guias/1','/Seguimiento/confirmaciones/1']){
        await navigate(route);
        if(!await evaluate("!!document.querySelector('h1') && !document.body.textContent.includes('Servicio temporalmente no disponible')"))throw new Error('Consulta falló '+route);
        for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]){
            await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
            if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento '+route+' '+label);
            const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
            const output=path.join(project,'entregables','preview',`seguimiento-${route.split('/').filter(Boolean).join('-')}-${label}.png`);
            await fs.mkdir(path.dirname(output),{recursive:true});await fs.writeFile(output,Buffer.from(screenshot.data,'base64'));
        }
        console.log('OK navegador: '+route+' escritorio y móvil sin desbordamiento.');
    }
    await navigate('/Presupuestos');
    await click('[data-document-open][href*="/guias/"]');
    await wait("document.querySelector('#document-dialog').open && document.querySelector('#document-dialog [name=\"Fields[NumeroActa]\"]')");
    if(!await evaluate("document.querySelector('#document-dialog [name=\"Fields[NumeroActa]\"]').value==='GUIA-0007B' && document.querySelector('#document-dialog .app-date-picker-display').value==='03/11/2026'"))throw new Error('Guía no recupera número/fecha de evento.');
    await click('#document-dialog [data-hyt-close]');
    await wait("!document.querySelector('#document-dialog').open");
    await navigate('/Seguimiento/oc-cliente');
    await click('[data-document-open][href*="confirmaciones"]');
    await wait("document.querySelector('#document-dialog [name=\"Fields[NumeroConfirmacion]\"]')");
    const operation=await evaluate("document.querySelector('#document-dialog [name=OperationId]').value");
    await evaluate("document.querySelector('#document-dialog [name=MotivoCambio]').value='Confirmación revisada';document.querySelector('#document-dialog [data-document-form]').requestSubmit()");
    await wait("!document.querySelector('#document-dialog').open && location.pathname==='/Seguimiento/oc-cliente'");
    console.log('OK navegador: guía desde presupuesto y confirmación desde OC con guardado popup y retorno. Operación presente: '+Boolean(operation));
    for(let attempt=0;attempt<3;attempt++){
        await click('[data-document-open][href*="/Editar"]');
        await wait("document.querySelector('#document-dialog [name=\"Fields[NumeroOCCliente]\"]')");
        if(!await evaluate("document.querySelector('#document-dialog [name=\"Fields[NumeroOCCliente]\"]').value==='OC-000A/01' && document.querySelectorAll('#document-dialog .app-date-picker-panel').length===2"))throw new Error('OC o paneles de fechas duplicados.');
        await click('#document-dialog [data-hyt-close]');
        await wait("!document.querySelector('#document-dialog').open");
    }
    console.log('OK navegador: tres aperturas OC sin duplicar paneles de calendario.');
    await navigate('/Presupuestos');await click('[data-document-open][href*="/guias/"]');
    await wait("document.querySelector('#document-dialog [name=Attachment]')");
    for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]){
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento popup '+label);
        const screenshot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview',`seguimiento-popup-guia-${label}.png`),Buffer.from(screenshot.data,'base64'));
    }
    const dom=await send('DOM.getDocument');const upload=await send('DOM.querySelector',{nodeId:dom.root.nodeId,selector:'#document-dialog input[name=Attachment]'});
    await send('DOM.setFileInputFiles',{nodeId:upload.nodeId,files:[path.join(project,'entregables','preview','anterior-una.pdf')]});
    await evaluate("document.querySelector('#document-dialog [name=MotivoCambio]').value='Adjuntar evidencia PDF';document.querySelector('#document-dialog [data-document-form]').requestSubmit()");
    await wait("location.pathname==='/Presupuestos' && !document.querySelector('#document-dialog').open");
    const fileCheck=await evaluate("fetch('/Seguimiento/guias/Archivos/100').then(async response=>({ok:response.ok,type:response.headers.get('content-type'),length:(await response.arrayBuffer()).byteLength}))");
    if(!fileCheck.ok||fileCheck.type!=='application/pdf'||fileCheck.length<1000)throw new Error('PDF privado no disponible tras carga.');
    await send('Emulation.setDeviceMetricsOverride',{width:1200,height:1000,deviceScaleFactor:1,mobile:false});
    await evaluate("document.body.replaceChildren(Object.assign(document.createElement('iframe'),{src:'/Seguimiento/guias/Archivos/100',title:'Evidencia PDF',style:'width:1100px;height:950px;border:0'}))");
    await delay(1500);
    const pdfScreenshot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview','seguimiento-pdf-privado.png'),Buffer.from(pdfScreenshot.data,'base64'));
    console.log('OK navegador: carga PDF desde popup y recuperación privada con tipo y contenido correctos.');
    await navigate('/Seguimiento/guias/1');
    await evaluate("const form=document.querySelector('form[action$=\"/Estado\"]');form.querySelector('[name=state]').value='ACEPTADA';form.querySelector('[name=eventDate]').value='2026-11-04';form.querySelector('[name=responsible]').value='Representante del cliente';form.requestSubmit()");
    await wait("document.body.textContent.includes('Cambio de estado registrado.')");
    await navigate('/Presupuestos/55');
    if(!await evaluate("document.querySelector('[name=NewState]')?.options.length===1 && document.querySelector('[name=NewState]').value==='ENVIADO' && !document.body.textContent.includes('Editar borrador')"))throw new Error('Regularización permite acciones ajenas al flujo.');
    await navigate('/Seguimiento/oc-cliente?confirmed=true');
    if(!await evaluate("document.querySelector('[name=confirmed]')?.value==='true' && document.body.textContent.includes('CONF-0001A')"))throw new Error('Filtro y confirmación no visibles.');
    console.log('OK navegador: aceptación auditada, regularización de emergencia y filtro de OC confirmadas.');
    await navigate('/Seguimiento/oc-cliente');await click('[data-document-open][href*="/Facturado"]');
    await wait("!!document.querySelector('#document-dialog [name=Series]')");
    await evaluate("document.querySelector('#document-dialog [name=Date]').value='2026-11-04';document.querySelector('#document-dialog [name=Series]').value='F001';document.querySelector('#document-dialog [name=Number]').value='000012A';document.querySelector('#document-dialog [data-document-form]').requestSubmit()");
    await wait("location.pathname==='/Seguimiento/oc-cliente'&&!document.querySelector('#document-dialog').open");
    console.log('OK navegador: captura y guardado Facturado desde popup OC conservando número alfanumérico.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
