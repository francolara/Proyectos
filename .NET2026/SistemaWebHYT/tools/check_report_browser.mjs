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
    await navigate('/Configuracion/HYT');
    if(!await evaluate("document.querySelector('[name=\"Fields[SubtituloCotizacion]\"]').value.includes('hytserviciosg@gmail.com - rene_torres35@hotmail.com')"))throw new Error('Subtítulo completo no está disponible en configuración.');
    const documentNode=await send('DOM.getDocument');const fileNode=await send('DOM.querySelector',{nodeId:documentNode.root.nodeId,selector:'#budget-signature'});
    await send('DOM.setFileInputFiles',{nodeId:fileNode.nodeId,files:[path.join(project,'GestionServicios.Web','wwwroot','images','logo-hyt.png')]});
    await evaluate("document.querySelector('#budget-signature').closest('form').requestSubmit()");
    await wait("document.querySelector('img[alt=\"Firma para presupuesto\"]')?.naturalWidth>0");
    await navigate('/Presupuestos/Nuevo');
    if(await evaluate("!!document.querySelector('.hyt-budget-print-card')"))throw new Error('Impresión visible antes de guardar.');
    for(const route of ['/Presupuestos/1','/Presupuestos/1/Editar']){
        await navigate(route);
        if(!await evaluate("(()=>{const header=document.querySelector('.hyt-budget-header'),card=header?.querySelector('.hyt-budget-print-card'),form=card?.querySelector('form');return card&&form.method==='post'&&form.target==='_blank'&&form.querySelector('[name=__RequestVerificationToken]')&&card.querySelectorAll('button[name=format]').length===2&&[...card.querySelectorAll('button')].every(e=>!e.disabled)&&!document.body.textContent.includes('Consultar impresiones conservadas');})()"))throw new Error('Tarjeta de impresión incorrecta en '+route);
    }
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate('/Presupuestos/1');
    await fs.mkdir(path.join(project,'entregables','preview','anterior-contacto-tienda'),{recursive:true});
    let headingShot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});await fs.writeFile(path.join(project,'entregables','preview','anterior-contacto-tienda','cabecera-impresion.png'),Buffer.from(headingShot.data,'base64'));
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
    if(!await evaluate("document.documentElement.scrollWidth<=innerWidth+1&&[...document.querySelectorAll('.hyt-budget-print-card button')].every(b=>b.getBoundingClientRect().width<=innerWidth)"))throw new Error('Tarjeta desborda en móvil.');
    headingShot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});await fs.writeFile(path.join(project,'entregables','preview','anterior-contacto-tienda','cabecera-impresion-movil.png'),Buffer.from(headingShot.data,'base64'));
    await send('Emulation.setDeviceMetricsOverride',{width:1280,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate('/Presupuestos/1');
    for(const format of ['NUEVO','ANTERIOR']){
        const prior=new Set((await send('Target.getTargets')).targetInfos.map(t=>t.targetId));
        await click('.hyt-budget-print-card button[value="'+format+'"]');
        let target;
        for(let i=0;i<100;i++){
            target=(await send('Target.getTargets')).targetInfos.find(t=>!prior.has(t.targetId)&&t.url.startsWith(fixture.url+'/Presupuestos/Impresiones/'));
            if(target)break;await delay(100);
        }
        if(!target)throw new Error('Botón no generó impresión: '+format);
        await send('Target.closeTarget',{targetId:target.targetId});
        console.log('OK botón real: '+format+' abre impresión en pestaña nueva.');
    }
    const output=path.join(project,'entregables','preview','anterior-contacto-tienda');await fs.mkdir(output,{recursive:true});
    const results=[];
    for(const [caseName,budget,format] of [['nuevo-una',1,'NUEVO'],['anterior-una',1,'ANTERIOR'],['nuevo-varias',50,'NUEVO'],['anterior-varias',50,'ANTERIOR'],['nuevo-vacio',51,'NUEVO'],['nuevo-usd',52,'NUEVO'],['nuevo-descripcion-larga',54,'NUEVO'],['anterior-vacio',51,'ANTERIOR'],['anterior-usd',52,'ANTERIOR'],['anterior-descripcion-larga',54,'ANTERIOR']]){
        await navigate('/Presupuestos/'+budget);
        const result=await evaluate(`(async()=>{
            const token=document.querySelector('input[name="__RequestVerificationToken"]').value;
            const response=await fetch('/Presupuestos/Impresiones/Generar',{method:'POST',body:new URLSearchParams({__RequestVerificationToken:token,budgetId:'${budget}',revisionId:'${budget}',version:btoa(String.fromCharCode(...new Uint8Array(8))),format:'${format}'})});
            return {url:response.url,status:response.status};
        })()`);
        if(result.status!==200||!result.url.includes('/Presupuestos/Impresiones/'))throw new Error('No se generó '+caseName+': '+JSON.stringify(result));
        await navigate(new URL(result.url).pathname);
        await evaluate('Promise.all([...document.images].map(image=>image.decode()))');
        const state=await evaluate(`({columns:document.querySelector('.report-detail thead tr').cells.length,rows:document.querySelectorAll('tbody tr').length,text:document.body.innerText,images:[...document.images].every(image=>image.complete&&image.naturalWidth>0)})`);
        if(state.columns!==(format==='NUEVO'?11:7)||!state.images||state.text.includes('#N/A'))throw new Error('Estructura incorrecta: '+caseName);
        if(caseName==='nuevo-usd'&&(!state.text.includes('USD')||!state.text.includes('$ 118.00')))throw new Error('Moneda o IGV incorrectos.');
        if(!await evaluate("(()=>{const p=document.querySelector('.new-services,.report-services');return p.textContent.includes('hytserviciosg@gmail.com - rene_torres35@hotmail.com') && getComputedStyle(p).textAlign==='center';})()"))throw new Error('Subtítulo configurado no está completo o centrado.');
        if(caseName==='nuevo-una'&&!await evaluate("[...document.querySelectorAll('.report-detail tbody tr')].some(row=>row.cells.length===11&&row.cells[4].textContent.trim()==='0001-A'&&row.cells[1].textContent.trim()==='F001')"))throw new Error('Posición o centro FICO no están mapeados.');
        if(caseName.endsWith('-una')&&state.text.includes('Servicio de prueba'))throw new Error('Se imprime la descripción interna del trabajo.');
        if(caseName==='nuevo-una'){
            const boxes=await evaluate(`(()=>{const cells=[...document.querySelectorAll('.new-zone td,.new-supplier td')],rects=cells.map(e=>e.getBoundingClientRect()),number=document.querySelector('.new-client tr:nth-child(2) td:nth-child(3)'),date=number.nextElementSibling,table=number.closest('table'),canvas=document.createElement('canvas'),ctx=canvas.getContext('2d');ctx.font=getComputedStyle(number).font;return {uniform:Math.max(...rects.map(r=>r.height))-Math.min(...rects.map(r=>r.height))<.2,wide:rects.every(r=>r.width>=38.5*96/25.4),center:cells.every(e=>getComputedStyle(e).textAlign==='center'),dateWidth:date.getBoundingClientRect().width/table.getBoundingClientRect().width,numberFits:ctx.measureText('202611000001').width+6<number.clientWidth};})()`);
            if(!boxes.uniform||!boxes.wide||!boxes.center||!boxes.numberFits||Math.abs(boxes.dateWidth-.093)>.005)throw new Error('Cabecera o cuadros incorrectos: '+JSON.stringify(boxes));
        }
        if(!await evaluate("(()=>{const p=document.querySelector('.new-services');return getComputedStyle(document.body).fontFamily.startsWith('Arial')&&getComputedStyle(document.querySelector('.report-detail')).fontFamily.startsWith('Arial')&&getComputedStyle(p).fontWeight==='400'&&getComputedStyle(p).borderTopColor==='rgb(201, 22, 32)';})()"))throw new Error('Tipografía o línea roja incorrecta.');
        if(format==='ANTERIOR'&&!await evaluate("(()=>{const t=document.querySelector('.report-old-table'),headers=[...t.tHead.rows[0].cells].map(c=>c.textContent.trim()),text=document.querySelector('.old-document').textContent;return headers.length===7&&headers[0]==='ITEM'&&headers[1]==='DESCRIPCIÓN'&&headers[6]==='TOTAL'&&getComputedStyle(t.tHead.rows[0].cells[0]).backgroundColor==='rgb(192, 0, 0)'&&text.includes('Coordinador/Encargado:')&&document.querySelector('[data-store-contact]').textContent.trim()==='Contacto de tienda'&&!text.includes('Registrado ·')&&!text.includes('Generado ')&&!text.includes('Los precios unitarios')&&text.split('Muy atentamente').length===2&&document.querySelector('.report-signature-image')?.naturalWidth>0;})()"))throw new Error('Formato anterior no cumple referencia o ajustes comunes.');
        if(caseName==='anterior-usd'&&!await evaluate("document.querySelector('.old-summary').textContent.includes('$ 100.00')&&!document.querySelector('.old-summary').textContent.includes('$ 18.00')&&!document.querySelector('.old-summary').textContent.includes('$ 118.00')"))throw new Error('Resumen neto del formato anterior incorrecto.');
        if(format==='ANTERIOR'&&!await evaluate("document.querySelector('.old-summary').rows.length===1&&document.querySelector('.old-context').rows[0].cells[1].textContent.trim()==='Tienda Centro'"))throw new Error('Resumen o nombre de tienda incorrectos.');
        const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
        if(caseName.endsWith('-una'))await fs.writeFile(path.join(output,caseName+'.png'),Buffer.from(screenshot.data,'base64'));
        const pdf=await send('Page.printToPDF',{printBackground:true,preferCSSPageSize:true});
        await fs.writeFile(path.join(output,caseName+'.pdf'),Buffer.from(pdf.data,'base64'));
        await fs.writeFile(path.join(output,caseName+'.html'),await evaluate('document.documentElement.outerHTML'));
        results.push({case:caseName,columns:state.columns,rows:state.rows});
        console.log('OK impresión Chrome: '+caseName+' / '+state.columns+' columnas.');
    }
    await fs.writeFile(path.join(output,'report-validation.json'),JSON.stringify(results,null,2));
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
