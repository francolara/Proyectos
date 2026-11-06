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
    for(const entity of ['tiendas','contactos']) {
        await navigate('/Maestros/'+entity);
        if(!await evaluate("document.body.textContent.includes('Código de cliente') && document.body.textContent.includes('CLI-000099') && document.body.textContent.includes('Cliente de prueba') && document.querySelector('a[href=\"/modulo/tiendas\"]') && document.querySelector('a[href=\"/modulo/contactos\"]')"))throw new Error('Faltan cliente o menús en '+entity);
        await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});await capture(entity+'-cliente-escritorio');
        await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
        if(!await evaluate('document.documentElement.scrollWidth<=innerWidth'))throw new Error('Listado desborda en móvil: '+entity);
        await capture(entity+'-cliente-movil');
    }
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    console.log('OK navegador: Tiendas y Contactos en menú; código y nombre del cliente en sus listados de escritorio/móvil.');
    await navigate('/Maestros/clientes/Nuevo');if(!await evaluate("document.querySelector('#internal-record-code').tagName==='OUTPUT' && !document.querySelector('.hyt-master-form [name=\"Fields[Codigo]\"]')"))throw new Error('Código cliente aún editable.');
    if(!await evaluate("document.querySelector('[name=\"Fields[CodigoSAP]\"]').maxLength===60 && !document.querySelector('[name=\"Fields[CodigoSAP]\"]').readOnly && document.querySelector('[data-field=CodigoSAP] label').textContent.includes('Código SAP')"))throw new Error('Código SAP no está disponible en Clientes.');
    await evaluate("document.querySelector('[name=\"Fields[TipoEntidad]\"]').value='JURIDICA';document.querySelector('[name=\"Fields[TipoEntidad]\"]').dispatchEvent(new Event('change'))");
    if(!await evaluate("!document.querySelector('[name=\"Fields[TipoIdentificacion]\"]') && !document.querySelector('[name=\"Fields[RazonSocial]\"]').disabled && document.querySelector('[name=\"Fields[Nombres]\"]').disabled && document.querySelector('[name=\"Fields[Apellidos]\"]').disabled"))throw new Error('Jurídica no distingue razón social y RUC.');
    await click('[data-client-link-add]');await click('[data-client-link-add]');
    await evaluate(`(()=>{const rows=document.querySelectorAll('[data-client-link-row]');rows[0].querySelector('[data-link-field="Description"]').value='Portal';rows[0].querySelector('[data-link-field="Url"]').value='https://example.com/portal';rows[1].querySelector('[data-link-field="Description"]').value='Facturas';rows[1].querySelector('[data-link-field="Url"]').value='javascript:alert(1)';rows[0].querySelector('[data-link-field="Url"]').dispatchEvent(new Event('input',{bubbles:true}));})()`);
    if(!await evaluate(`(()=>{const rows=document.querySelectorAll('[data-client-link-row]');return !rows[0].querySelector('[data-client-link-open]').hidden && rows[1].querySelector('[data-client-link-open]').hidden;})()`))throw new Error('Iconos permiten un esquema inseguro o no activan HTTPS.');
    await evaluate(`(()=>{const input=document.querySelectorAll('[data-client-link-row]')[1].querySelector('[data-link-field="Url"]');input.value='https://example.com/facturas';input.dispatchEvent(new Event('input',{bubbles:true}));})()`);
    await click('[data-client-link-remove]');await click('[data-client-link-add]');
    await evaluate(`(()=>{const row=document.querySelector('[data-client-link-rows]').lastElementChild;row.querySelector('[data-link-field="Description"]').value='Documentos adicionales';row.querySelector('[data-link-field="Url"]').value='https://example.com/adicional';row.querySelector('[data-link-field="Url"]').dispatchEvent(new Event('input',{bubbles:true}));})()`);
    if(!await evaluate(`document.querySelector('[name="DocumentLinks[0].Description"]').value==='Facturas' && document.querySelector('[name="DocumentLinks[1].Description"]').value==='Documentos adicionales'`))throw new Error('No renumera enlaces al quitar filas.');
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});await capture('cliente-juridica-escritorio');
    await evaluate("document.querySelector('[name=\"Fields[TipoEntidad]\"]').value='NATURAL';document.querySelector('[name=\"Fields[TipoEntidad]\"]').dispatchEvent(new Event('change'))");
    if(!await evaluate("document.querySelector('[name=\"Fields[RazonSocial]\"]').disabled && !document.querySelector('[name=\"Fields[Nombres]\"]').disabled && !document.querySelector('[name=\"Fields[Apellidos]\"]').disabled && !document.querySelector('[name=\"Fields[TipoIdentificacion]\"]') && document.querySelector('[data-field=NumeroIdentificacion] label').textContent.includes('CE')"))throw new Error('Natural no habilita nombres y DNI/CE.');
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});await capture('cliente-natural-movil');
    if(!await evaluate('document.documentElement.scrollWidth<=innerWidth'))throw new Error('Cliente móvil desborda.');
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    console.log('OK navegador: jurídica habilita razón social/RUC y natural nombres/apellidos; tipo de documento ausente sin validación.');
    await evaluate(`document.querySelector('[name="Fields[Nombres]"]').value='Ana';document.querySelector('[name="Fields[Apellidos]"]').value='Pérez';document.querySelector('[name="Fields[NumeroIdentificacion]"]').value='12345678';document.querySelector('.hyt-master-form').requestSubmit()`);
    await wait("location.pathname==='/Maestros/clientes/3' && document.readyState==='complete'");
    if(!await evaluate(`fetch('/__tests/maintenance-state').then(r=>r.json()).then(data=>{const payload=JSON.parse(data.payload);return payload.Enlaces.length===2&&payload.Enlaces[0].Descripcion==='Facturas'&&payload.Enlaces[1].Url==='https://example.com/adicional';})`))throw new Error('Los enlaces no llegaron al guardado del cliente.');
    if(!await evaluate(`document.querySelectorAll('[data-client-link-open][href]').length===2 && [...document.querySelectorAll('[data-client-link-open][href]')].every(a=>a.target==='_blank'&&a.rel.includes('noopener'))`))throw new Error('La consulta no permite abrir enlaces en pestaña segura.');
    await capture('cliente-enlaces-consulta');
    console.log('OK navegador: varios enlaces, descripción, icono seguro, retiro/reindexado y guardado con cliente.');
    await navigate('/Maestros/unidades/Nuevo');if(!await evaluate("document.querySelector('#internal-record-code').tagName==='OUTPUT' && document.body.textContent.includes('Abreviatura') && !document.querySelector('[name=\"Fields[Codigo]\"]').readOnly"))throw new Error('Unidad no separa identificador y abreviatura.');
    await navigate('/Maestros/materiales/1/Editar');
    if(!await evaluate("!document.querySelector('[name=\"Fields[Posicion]\"]') && document.querySelector('#internal-record-code').tagName==='OUTPUT' && document.querySelector('a[href=\"/modulo/especialidades\"]') && document.querySelector('a[href=\"/modulo/tipos-servicio\"]')"))throw new Error('Maestro o menú no ajustado.');
    await evaluate("document.querySelector('[name=\"Fields[Descripcion]\"]').value='Cambio del material todavía sin guardar'");
    for(const [name,width,height]of[['escritorio',1440,1000],['movil',390,844]]){await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Ficha material desborda '+name);await capture('material-tarifas-'+name);}
    await click('[data-material-price-open][href*="/Nuevo"]');await frameReady('[name="PricePEN"]');
    if(!await evaluate(`(()=>{const doc=${frameDoc};return document.querySelector('#material-price-title').textContent==='Agregar tarifario'&&!doc.body.textContent.toLowerCase().includes('versión')&&!doc.body.textContent.toLowerCase().includes('partidas')&&[...doc.querySelectorAll('input:not([type=hidden]),select')].map(e=>e.name).sort().join(',')==='Active,PricePEN,PriceUSD,TaxCode';})()`))throw new Error('La ventana sigue mostrando campos técnicos.');
    await evaluate(`(()=>{const doc=${frameDoc};doc.querySelector('[name="TaxCode"]').value='CON_IGV';doc.querySelector('[name="PricePEN"]').value='12.34';doc.querySelector('[name="PriceUSD"]').value='';})()`);
    const picture=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});await fs.writeFile(path.join(project,'entregables','preview','agregar-tarifario-simple.png'),Buffer.from(picture.data,'base64'));
    await evaluate(`${frameDoc}.querySelector('form').requestSubmit()`);
    await wait("!document.querySelector('#material-price-dialog').open && document.querySelector('[data-material-price-status]').textContent==='Precios actualizados.'");
    if(!await evaluate("document.querySelector('[name=\"Fields[Descripcion]\"]').value==='Cambio del material todavía sin guardar'"))throw new Error('Agregar tarifario pierde cambios del material.');
    if(!await evaluate("fetch('/__tests/maintenance-state').then(r=>r.json()).then(data=>JSON.parse(data.payload).MaterialId===1 && JSON.parse(data.payload).PrecioPEN===12.34 && JSON.parse(data.payload).PrecioUSD===null)"))throw new Error('Tarifario no llega al repositorio.');
    await click('[data-material-price-open][href*="/Editar"]');await frameReady('[name="PricePEN"]');
    await evaluate(`(()=>{const doc=${frameDoc};doc.querySelector('[name="PricePEN"]').value='88.22';doc.querySelector('[name="PriceUSD"]').value='20.00';doc.querySelector('[name="Active"]').value='false';doc.querySelector('form').requestSubmit();})()`);await wait("!document.querySelector('#material-price-dialog').open");
    if(!await evaluate("fetch('/__tests/maintenance-state').then(r=>r.json()).then(data=>JSON.parse(data.payload).PrecioPEN===88.22 && JSON.parse(data.payload).PrecioUSD===20 && JSON.parse(data.payload).Activo===false)"))throw new Error('Edición/estado del tarifario no llegan al repositorio.');
    console.log('OK navegador: Agregar tarifario muestra solo IGV, PEN/USD y estado; guarda sin pasos técnicos y conserva campos del material.');
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await send('Page.navigate',{url:fixture.url+'/Maestros/tarifarios'});await wait("location.pathname==='/Tarifarios' && document.readyState==='complete'");
    if(!await evaluate("location.pathname==='/Tarifarios' && document.body.textContent.includes('MAT-000001') && document.body.textContent.includes('Material de prueba')"))throw new Error('Menú Tarifarios no muestra los precios del material.');
    await capture('tarifarios-unificados-escritorio');
    await click('a[href="/Tarifarios/Nuevo"]');await wait("document.querySelector('[name=selectedMaterialId]')");
    await click('[data-open-lookup]');await wait("document.querySelector('#master-lookup').open && document.querySelector('[data-lookup-rows] button')");
    await click('[data-lookup-rows] button');await wait("document.querySelector('[name=selectedMaterialId]').value==='1'");
    await evaluate("document.querySelector('[name=PricePEN]').value='45.67';document.querySelector('[name=PriceUSD]').value='12.34';document.querySelector('[name=TaxCode]').value='CON_IGV'");
    await capture('tarifario-nuevo-desde-menu');
    await evaluate("document.querySelector('form[data-hyt-save]').requestSubmit()");await wait("location.pathname==='/Tarifarios'");
    if(!await evaluate("fetch('/__tests/maintenance-state').then(r=>r.json()).then(data=>JSON.parse(data.payload).PrecioPEN===45.67 && JSON.parse(data.payload).PrecioUSD===12.34 && JSON.parse(data.payload).MaterialId===1)"))throw new Error('Alta desde menú no usa el registro compartido.');
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
    if(!await evaluate('document.documentElement.scrollWidth<=innerWidth+1'))throw new Error('Tarifarios desborda en móvil.');
    await capture('tarifarios-unificados-movil');
    console.log('OK navegador: listado unificado y alta desde menú con selección popup de material, IGV y precios; móvil sin desborde.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
