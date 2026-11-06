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

    const results=[];
    function check(value,label){if(!value)throw new Error(label);results.push(label);}
    async function contrast(selector){return await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)return 0;const cs=getComputedStyle(e);let p=e,bg;while(p){const c=getComputedStyle(p).backgroundColor;if(c!=='rgba(0, 0, 0, 0)'&&c!=='transparent'){bg=c;break;}p=p.parentElement;}const values=c=>(c.match(/[0-9.]+/g)||[]).slice(0,3).map(Number);const lum=c=>values(c).map(v=>v/255).map(v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4).reduce((a,v,i)=>a+v*[.2126,.7152,.0722][i],0);const a=lum(cs.color),b=lum(bg||'rgb(255,255,255)');return (Math.max(a,b)+.05)/(Math.min(a,b)+.05);})()`);}
    await send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
    await send('Emulation.setEmulatedMedia',{features:[{name:'prefers-color-scheme',value:'dark'}]});
    await navigate('/Acceso');await evaluate("localStorage.removeItem('hyt.theme')");await navigate('/Acceso');
    check(await evaluate("document.documentElement.dataset.hytTheme==='dark'"),'Tema inicial sigue preferencia del sistema');
    check(await contrast('.hyt-auth-form label')>=4.5,'Acceso legible en modo oscuro');
    await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('form').requestSubmit()`);await wait("location.pathname==='/'");
    await navigate('/OCCliente/1/Facturado');
    check(await evaluate(`(async()=>{const form=document.querySelector('[data-document-form]'),data=new FormData(form);data.set('Date','2026-11-05');data.set('Series','F001');data.set('Number','0001-UI');if(data.get('ReferenceId'))data.set('Reason','Verificación de interfaz');const response=await fetch(form.action,{method:'POST',body:data});return response.ok;})()`),'Preparación de referencia existente para comprobar responsables de Facturado');
    const listRoutes=['/Maestros/clientes','/Maestros/tiendas','/Maestros/contactos','/Maestros/materiales','/Presupuestos','/Seguimiento/oc-cliente','/Seguimiento/guias','/Configuracion/TasasIGV','/Reportes','/Presupuestos/Impresiones?budgetId=1&revisionId=1'];
    for(const theme of ['light','dark']){
        await navigate('/Maestros/clientes');if(await evaluate('document.documentElement.dataset.hytTheme')!==theme)await click('[data-hyt-theme-toggle]');
        check(await evaluate(`localStorage.getItem('hyt.theme')===${JSON.stringify(theme)}`),'Preferencia guardada '+theme);
        for(const route of listRoutes){
            await navigate(route);
            check(await evaluate(`document.documentElement.dataset.hytTheme===${JSON.stringify(theme)} && document.querySelector('.hyt-workspace-wide')?.clientWidth>1600`),'Listado ancho y tema persistente '+theme+' '+route);
            check(await evaluate("getComputedStyle(document.querySelector('.hyt-page-header')).borderLeftWidth==='5px'"),'Título destacado '+theme+' '+route);
            check(await contrast('.hyt-page-header h1')>=4.5,'Contraste de título '+theme+' '+route);
            check(await evaluate("[...document.querySelectorAll('.hyt-table a,.hyt-grid-actions a')].every(a=>getComputedStyle(a).textDecorationLine==='none' && getComputedStyle(a).borderTopStyle==='solid')"),'Acciones con aspecto de botón '+theme+' '+route);
            check(await evaluate('document.documentElement.scrollWidth<=innerWidth'),'Sin desbordamiento de página '+theme+' '+route);
            if(route==='/Maestros/clientes')await capture('interfaz-listado-'+theme);
        }
        for(const route of ['/Maestros/clientes/1/Editar','/Maestros/materiales/1/Editar','/Presupuestos/1/Editar','/Configuracion/HYT','/Seguimiento/guias/1/Editar','/OCCliente/1/Facturado']){
            await navigate(route);
            check(await evaluate("!document.querySelector('.hyt-workspace-wide') && document.querySelector('main').clientWidth<=1440"),'Formulario conserva ancho '+theme+' '+route);
            check(await evaluate("!document.body.textContent.includes('Historial de cambios') && !document.body.textContent.includes('Historial de acciones') && !document.body.textContent.includes('Historial de la revisión y del expediente')"),'Sin historial general en edición '+theme+' '+route);
            if(await evaluate("!!document.querySelector('#internal-record-code')"))check(await contrast('#internal-record-code')>=4.5,'Código interno legible '+theme+' '+route);
            check(await evaluate("!!document.querySelector('.hyt-trace-summary')"),'Edición existente muestra fechas y responsables '+theme+' '+route);
            if(await evaluate("!!document.querySelector('.hyt-trace-summary')")){
                check(await evaluate("getComputedStyle(document.querySelector('.hyt-trace-summary dl')).gridTemplateColumns.split(' ').length===4"),'Cuatro columnas de responsables '+theme+' '+route);
                check(await contrast('.hyt-trace-summary dd')>=4.5,'Contraste de responsables '+theme+' '+route);
            }
            if(route.includes('/materiales/')){
                check(await evaluate(`(()=>{const e=document.querySelector('[name="Fields[Descripcion]"]'),p=e.closest('.hyt-master-fields');return e.tagName==='TEXTAREA'&&e.rows===3&&e.maxLength===1500&&e.getBoundingClientRect().width>p.clientWidth*.95&&e.clientHeight>=90;})()`),'Material: descripción de tres líneas a todo el ancho '+theme);
                check(await contrast('[name="Fields[Descripcion]"]')>=4.5,'Contraste del área de descripción '+theme);
                await capture('interfaz-material-'+theme);
                await click('[data-reference="UnidadMedidaId"] [data-open-lookup]');await wait("document.querySelector('#master-lookup').getAttribute('aria-busy')==='false'");
                check(await evaluate("getComputedStyle(document.querySelector('#master-lookup .hyt-modal-header')).backgroundColor==='rgb(201, 22, 32)'"),'Barra roja de ayuda '+theme);
                check(await contrast('#master-lookup h2')>=4.5,'Contraste de título de popup '+theme);
                await capture('interfaz-ayuda-'+theme);
                await send('Input.dispatchKeyEvent',{type:'keyDown',key:'Escape',code:'Escape',windowsVirtualKeyCode:27});await send('Input.dispatchKeyEvent',{type:'keyUp',key:'Escape',code:'Escape',windowsVirtualKeyCode:27});
                check(await evaluate("!document.querySelector('#master-lookup').open && document.activeElement.matches('[data-open-lookup]')"),'Escape cierra y devuelve foco '+theme);
                await click('[data-material-price-open]');await frameReady('[name="PricePEN"]');
                check(await evaluate(`document.querySelector('[data-material-price-frame]').contentDocument.documentElement.dataset.hytTheme===${JSON.stringify(theme)}`),'Tema en iframe de tarifas '+theme);
                check(await evaluate("document.querySelector('[data-material-price-frame]').contentDocument.querySelector('h1').getBoundingClientRect().width<=1 && document.querySelector('#material-price-title').textContent==='Agregar tarifario'"),'Un título visible para popup de tarifas '+theme);
                await evaluate("document.querySelector('[data-hyt-theme-toggle]').click()");
                await wait(`document.querySelector('[data-material-price-frame]').contentDocument.documentElement.dataset.hytTheme!==${JSON.stringify(theme)}`);
                check(true,'Tema sincroniza con popup abierto '+theme);await evaluate("document.querySelector('[data-hyt-theme-toggle]').click()");
                await click('[data-material-price-close]');
            }
        }
        await navigate('/Presupuestos');await click('.app-date-picker-trigger');await wait("!!document.querySelector('.app-date-picker-panel:not([hidden])')");
        check(await evaluate("document.querySelector('.app-date-picker-panel:not([hidden]) .app-date-picker-banner').textContent==='Seleccionar fecha'"),'Calendario con cabecera '+theme);
        check(await contrast('.app-date-picker-panel:not([hidden]) .app-date-picker-title strong')>=4.5,'Calendario legible '+theme);await capture('interfaz-calendario-'+theme);
        await navigate('/Seguimiento/oc-cliente');await click('[data-document-open]');await wait("document.querySelector('#document-dialog').open && !!document.querySelector('#document-dialog [data-document-form]')");
        check(await evaluate("getComputedStyle(document.querySelector('#document-dialog .hyt-modal-header')).backgroundColor==='rgb(201, 22, 32)'"),'Popup documental con cabecera HYT '+theme);
        check(await evaluate("!document.querySelector('#document-dialog').textContent.includes('Historial de acciones')"),'Popup documental sin historial de edición '+theme);
        await capture('interfaz-documento-'+theme);await click('[data-hyt-close]');
        await navigate('/');await wait("!!document.querySelector('[data-chart-series]')");
        check(await contrast('.hyt-dashboard-metrics strong')>=4.5,'Indicadores legibles '+theme);
        check(await evaluate("!document.querySelector('[data-chart-link]') || !document.querySelector('[data-chart-link]').hidden || getComputedStyle(document.querySelector('[data-chart-link]')).display==='none'"),'Botones ocultos no muestran contornos vacíos '+theme);await capture('interfaz-dashboard-'+theme);
        for(const width of [768,390]){
            await send('Emulation.setDeviceMetricsOverride',{width,height:950,deviceScaleFactor:1,mobile:width===390});
            for(const route of ['/Maestros/clientes','/Maestros/materiales/1/Editar','/Presupuestos','/Seguimiento/oc-cliente','/Configuracion/HYT','/Reportes']){
                await navigate(route);
                check(await evaluate('document.documentElement.scrollWidth<=innerWidth'),'Adaptable '+theme+' '+width+' '+route);
                check(await evaluate("(()=>{const r=document.querySelector('[data-hyt-theme-toggle]').getBoundingClientRect();return r.width>0&&r.height>=44;})()"),'Selector de tema visible en móvil/tablet '+theme+' '+width+' '+route);
                if(width===390&&await evaluate("!!document.querySelector('.hyt-trace-summary')"))check(await evaluate("getComputedStyle(document.querySelector('.hyt-trace-summary dl')).gridTemplateColumns.split(' ').length===1"),'Responsables en una columna móvil '+theme+' '+route);
                if(route==='/Maestros/materiales/1/Editar')await capture('interfaz-material-'+theme+'-'+width);
            }
            await navigate('/Presupuestos');await click('.app-date-picker-trigger');await wait("!!document.querySelector('.app-date-picker-panel:not([hidden])')");
            check(await evaluate("(()=>{const r=document.querySelector('.app-date-picker-panel:not([hidden])').getBoundingClientRect();return r.left>=0&&r.right<=innerWidth&&r.top>=0&&r.bottom<=innerHeight;})()"),'Calendario dentro de ventana '+theme+' '+width);
        }
        await send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
    }
    await navigate('/Maestros/materiales/1/Editar');
    const sampleDescription=['Primera línea','Segunda línea','Tercera línea'].join(String.fromCharCode(10));
    await evaluate(`document.querySelector('[name="Fields[Descripcion]"]').value=${JSON.stringify(sampleDescription)};document.querySelector('.hyt-master-form').requestSubmit()`);
    await wait("location.pathname==='/Maestros/materiales/1' && document.readyState==='complete'");
    check(await evaluate(`fetch('/__tests/maintenance-state').then(r=>r.json()).then(data=>JSON.parse(data.payload).Descripcion.split(String.fromCharCode(13)).join('')===${JSON.stringify(sampleDescription)})`),'Guardado conserva saltos de línea del material');
    await navigate('/Reportes');await send('Emulation.setEmulatedMedia',{media:'print'});
    check(await evaluate("getComputedStyle(document.body).backgroundColor==='rgb(255, 255, 255)'"),'Reporte impreso con fondo blanco aunque el tema sea oscuro');
    await send('Emulation.setEmulatedMedia',{media:'',features:[{name:'prefers-color-scheme',value:'dark'}]});
    await fs.writeFile(path.join(project,'entregables','preview','verificacion-interfaz.json'),JSON.stringify({date:'2026-11-05',checks:results.length,results},null,2));
    console.log(`OK interfaz: ${results.length} comprobaciones en ambos temas, escritorio/tablet/móvil, popups y contraste.`);
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
