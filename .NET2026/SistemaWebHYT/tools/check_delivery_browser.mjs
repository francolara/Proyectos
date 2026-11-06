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

    await navigate('/Acceso');
    await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('form').requestSubmit()`);
    await wait("location.pathname==='/'");
    const runtimeErrors=[];socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.method==='Runtime.exceptionThrown')runtimeErrors.push(message.params.exceptionDetails.text);});await send('Runtime.enable');
    const token=await evaluate("fetch('/OCCliente/1/Facturado').then(response=>response.text()).then(html=>new DOMParser().parseFromString(html,'text/html').querySelector('[name=__RequestVerificationToken]').value)");
    async function configure(options){const result=await evaluate(`fetch('/__tests/reporting',{method:'POST',body:new URLSearchParams({...${JSON.stringify(options)},__RequestVerificationToken:${JSON.stringify(token)}})}).then(response=>response.status)`);if(result!==200)throw new Error('No se configuró el caso QA.');}
    await configure({empty:'true',rows:'0'});await navigate('/');
    await wait("document.querySelector('[data-chart-note]').textContent.includes('No hay registros')");
    if(!await evaluate("document.querySelector('[data-chart=trend]').style.display==='none' && !document.body.textContent.includes('NaN') && document.body.textContent.includes('Sin importes en este período')"))throw new Error('Vacío dashboard incorrecto.');
    await configure({empty:'false',budgets:'true',orders:'false'});await navigate('/');
    await wait("document.querySelectorAll('[data-chart=trend] path').length===1");
    if(!await evaluate("!document.querySelector('[data-chart=donut]') && !document.querySelector('[data-chart-series=Ordenes]')"))throw new Error('Fuente OC oculta aparece en gráficos.');
    await configure({empty:'false',budgets:'false',orders:'false'});await navigate('/');
    if(!await evaluate("document.body.textContent.includes('Tu perfil no tiene acceso a las fuentes') && !document.querySelector('[data-chart=trend]')"))throw new Error('Estado sin fuentes no se identifica.');
    await configure({empty:'false',budgets:'true',orders:'true',rows:'20'});
    for(const kind of ['GANADOS','OC_COMPLETADAS','OC_CON_ACTA','OC_PENDIENTES_FACTURA']){
        await navigate('/Reportes?kind='+kind);
        if(!await evaluate("document.querySelectorAll('tbody tr').length===20 && document.body.textContent.includes('9,676.00') && !document.querySelector('main script:not([src])')"))throw new Error('Reporte incompleto '+kind);
        for(const [label,width,height] of [['escritorio',1440,1000],['movil',390,844]]){
            await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
            if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Reporte desborda '+kind+' '+label);
            if(kind==='GANADOS'){const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});await fs.writeFile(path.join(project,'entregables','preview',`reporte-ganados-${label}.png`),Buffer.from(screenshot.data,'base64'));}
        }
        const pdf=await send('Page.printToPDF',{printBackground:true,preferCSSPageSize:true});await fs.writeFile(path.join(project,'entregables','preview','reporte-'+kind.toLowerCase()+'.pdf'),Buffer.from(pdf.data,'base64'));
        console.log('OK QA: '+kind+' veinte filas, moneda, móvil e impresión HTML.');
    }
    await configure({rows:'0'});await navigate('/Reportes');if(!await evaluate("document.body.textContent.includes('Sin registros para los filtros')"))throw new Error('Reporte vacío incorrecto.');
    await navigate('/Seguimiento/oc-cliente');await click('[data-document-open][href*="/Facturado"]');await wait("document.querySelector('#document-dialog [name=Series]')");
    await evaluate("const form=document.querySelector('#document-dialog form');form.querySelector('[name=Date]').value='2026-11-04';form.querySelector('[name=Series]').value='F001';form.querySelector('[name=Number]').value='000012A';form.requestSubmit()");
    await wait("!document.querySelector('#document-dialog').open");
    await click('[data-document-open][href*="/Facturado"]');await wait("document.querySelector('#document-dialog [name=Reason]')");
    if(!await evaluate("document.querySelector('#document-dialog [name=Number]').value==='000012A'"))throw new Error('No recupera referencia Facturado.');
    await evaluate("const form=document.querySelector('#document-dialog form');form.querySelector('[name=Number]').value='000099B';form.querySelector('[name=Reason]').value='Corrección de número externo';form.requestSubmit()");await wait("!document.querySelector('#document-dialog').open");
    await click('[data-document-open][href*="/Facturado"]');await wait("document.querySelector('#document-dialog [name=Number]')?.value==='000099B'");
    if(!await evaluate("document.querySelector('#document-dialog').textContent.includes('000012A')"))throw new Error('Antecedente de corrección no visible.');
    await evaluate("const form=document.querySelector('#document-dialog form');form.querySelector('[name=Withdraw]').checked=true;form.querySelector('[name=Reason]').value='Retiro por anulación externa';form.requestSubmit()");await wait("!document.querySelector('#document-dialog').open");
    await click('[data-document-open][href*="/Facturado"]');await wait("document.querySelector('#document-dialog [name=ReferenceId]')?.value===''");
    if(!await evaluate("document.querySelector('#document-dialog').textContent.includes('Retirada') && document.querySelector('#document-dialog').textContent.includes('Retiro por anulación externa')"))throw new Error('Retiro no conserva historia visible.');
    if(runtimeErrors.length)throw new Error('Errores JavaScript: '+runtimeErrors.join(';'));
    console.log('OK QA: dashboard vacío/fuentes parciales, reportes vacíos y Facturado con corrección/retiro conservando historial.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
