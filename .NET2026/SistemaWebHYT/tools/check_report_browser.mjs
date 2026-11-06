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
    const output=path.join(project,'entregables','preview');await fs.mkdir(output,{recursive:true});
    const results=[];
    for(const [caseName,budget,format] of [['nuevo-una',1,'NUEVO'],['anterior-una',1,'ANTERIOR'],['nuevo-varias',50,'NUEVO'],['anterior-varias',50,'ANTERIOR'],['nuevo-vacio',51,'NUEVO'],['nuevo-usd',52,'NUEVO'],['nuevo-descripcion-larga',54,'NUEVO']]){
        await navigate('/Presupuestos/'+budget);
        const result=await evaluate(`(async()=>{
            const token=document.querySelector('input[name="__RequestVerificationToken"]').value;
            const response=await fetch('/Presupuestos/Impresiones/Generar',{method:'POST',body:new URLSearchParams({__RequestVerificationToken:token,budgetId:'${budget}',revisionId:'${budget}',version:btoa(String.fromCharCode(...new Uint8Array(8))),format:'${format}'})});
            return {url:response.url,status:response.status};
        })()`);
        if(result.status!==200||!result.url.includes('/Presupuestos/Impresiones/'))throw new Error('No se generó '+caseName+': '+JSON.stringify(result));
        await navigate(new URL(result.url).pathname);
        await evaluate('Promise.all([...document.images].map(image=>image.decode()))');
        const state=await evaluate(`({columns:document.querySelector('thead tr').cells.length,rows:document.querySelectorAll('tbody tr').length,text:document.body.innerText,images:[...document.images].every(image=>image.complete&&image.naturalWidth>0)})`);
        if(state.columns!==(format==='NUEVO'?11:7)||!state.images||state.text.includes('#N/A'))throw new Error('Estructura incorrecta: '+caseName);
        if(caseName==='nuevo-usd'&&(!state.text.includes('USD')||!state.text.includes('$ 118.00')||!state.text.includes('incluyen IGV')))throw new Error('Moneda o IGV incorrectos.');
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
