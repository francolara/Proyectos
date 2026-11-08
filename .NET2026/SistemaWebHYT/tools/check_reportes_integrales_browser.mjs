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
    const errors=[];
    await send('Runtime.enable');
    for(const route of ['/Reportes?kind=PRESUPUESTO','/Reportes?kind=OC_CLIENTE','/']){
      await navigate(route);
      if(!await evaluate(`!document.querySelector('#report-situation')&&!document.querySelector('[name=client]')&&!document.querySelector('[name=store]')`))throw new Error('Filtros comerciales antiguos');
      for(const [width,height] of [[1600,1100],[390,844]]){
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Desbordamiento '+route+' '+width);
        const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});const name=route==='/'?'dashboard':route.includes('OC_CLIENTE')?'reporte-oc':'reporte-presupuesto';await fs.writeFile(path.join(project,'entregables','preview',`${name}-integral-${width}.png`),Buffer.from(shot.data,'base64'));
      }
      await send('Emulation.setDeviceMetricsOverride',{width:1600,height:1100,deviceScaleFactor:1,mobile:false});
    }
    if(!await evaluate(`document.querySelectorAll('[data-state-source] rect').length>0&&document.querySelector('[data-chart=trend] path')&&document.querySelector('[data-chart=donut] [data-segment]')`))throw new Error('Gráficos del nuevo panel no dibujados');
    await navigate('/Reportes?kind=PRESUPUESTO');
    await evaluate(`document.querySelector('#report-kind').value='OC_CLIENTE';document.querySelector('#report-kind').dispatchEvent(new Event('change'))`);
    await wait(`location.search.includes('kind=OC_CLIENTE')&&!!document.querySelector('option[value=TRABAJO_COMPLETADO]')`);
    if(!await evaluate(`document.querySelector('#report-kind').options.length===2&&!document.querySelector('option[value=APROBADO]')`))throw new Error('Catálogo o estados no cambian por fuente');
    await navigate('/Reportes/Imprimir?kind=PRESUPUESTO&type=EMERGENCIA&state=APROBADO&search=prueba');
    if(!await evaluate(`document.querySelectorAll('tbody tr').length===41&&document.querySelector('.hyt-report-brand img').naturalWidth>0&&document.body.textContent.includes('Estado: Aceptada')`))throw new Error('Impresión incompleta o logo ausente');
    await send('Emulation.setDeviceMetricsOverride',{width:1600,height:1100,deviceScaleFactor:1,mobile:false});
    const pdf=await send('Page.printToPDF',{printBackground:true,preferCSSPageSize:true});await fs.writeFile(path.join(project,'entregables','preview','reporte-integral-horizontal.pdf'),Buffer.from(pdf.data,'base64'));
    const shot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview','reporte-integral-impresion.png'),Buffer.from(shot.data,'base64'));
    console.log('OK: dos reportes, filtros dinámicos, referencias, impresión completa 41 filas con logo y PDF horizontal; dashboard con curvas/columnas/dona; escritorio y móvil.');
    await navigate('/');await click('[data-hyt-theme-toggle]');await wait(`document.documentElement.dataset.hytTheme==='dark'`);
    if(!await evaluate(`getComputedStyle(document.querySelector('.hyt-dashboard-states a')).backgroundColor!=='rgb(255, 255, 255)'`))throw new Error('Estado sin tema oscuro');
    await navigate('/Reportes?kind=OC_CLIENTE');
    if(!await evaluate(`document.documentElement.dataset.hytTheme==='dark'&&getComputedStyle(document.querySelector('.hyt-report-summary')).backgroundColor!=='rgb(255, 255, 255)'`))throw new Error('Reporte sin tema oscuro');
    const darkShot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview','reporte-integral-oscuro.png'),Buffer.from(darkShot.data,'base64'));
    await navigate('/');
    if(!await evaluate(`document.querySelector('[data-chart-series=Confirmaciones]').closest('label').textContent.trim()==='Confirmación del cliente'&&document.querySelector('[data-chart-series=Marcados]').closest('label').textContent.trim()==='Facturado'&&!document.body.textContent.includes('Primera confirmación')&&!document.body.textContent.includes('Primer marcado')`))throw new Error('Nombres del dashboard no corresponden al flujo actual');
    await click('[data-chart-series=Confirmaciones]');await click('[data-chart-series=Marcados]');
    await wait(`document.querySelector('[data-chart-readout]').textContent.includes('Confirmación del cliente:')&&document.querySelector('[data-chart-readout]').textContent.includes('Facturado:')`);
    const namesShot=await send('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(project,'entregables','preview','dashboard-confirmacion-facturado.png'),Buffer.from(namesShot.data,'base64'));
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
