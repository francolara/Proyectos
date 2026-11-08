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
    async function wait(expression,attempts=100) {
        for(let i=0;i<attempts;i++){ if(await evaluate("!!document.body && ("+expression+")")) return; await delay(100); }
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


    const errors=[];socket.addEventListener('message',event=>{const item=JSON.parse(event.data);if(item.method==='Runtime.exceptionThrown')errors.push(item.params.exceptionDetails.text);});
    await send('Runtime.enable');await send('Network.enable');await send('Network.setBlockedURLs',{urls:['*challenges.cloudflare.com/*']});
    const visible="document.querySelector('#hyt-loading') && !document.querySelector('#hyt-loading').hidden";
    const idle="globalThis.hytLoading?.pending===0 && (!document.querySelector('#hyt-loading') || document.querySelector('#hyt-loading').hidden)";
    async function capture(name){const picture=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});await fs.writeFile(path.join(project,'entregables','preview',name+'.png'),Buffer.from(picture.data,'base64'));}
    await navigate('/Acceso');await wait("globalThis.hytLoading && !document.querySelector('[data-manual-captcha]').hidden && document.querySelector('[data-captcha-image]').naturalWidth===260");
    if(!await evaluate("!document.querySelector('button[type=submit]').disabled && document.querySelector('#CaptchaAnswer').required"))throw new Error('Falta captcha inicial sin claves.');
    for(const [name,width,height]of[['escritorio',1440,1000],['movil',390,844]]){
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('new Promise(requestAnimationFrame)');
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Acceso manual desborda en '+name);
        await capture('captcha-manual-'+name);
    }
    const firstId=await evaluate("document.querySelector('#CaptchaId').value");await click('[data-captcha-refresh]');
    await wait(`document.querySelector('#CaptchaId').value!==${JSON.stringify(firstId)} && !document.querySelector('[data-captcha-refresh]').disabled && document.querySelector('[data-captcha-image]').naturalWidth===260`);
    await evaluate("fetch('/__tests/challenge-mode?mode=widget').then(r=>r.json())");await navigate('/Acceso');
    await wait("document.querySelector('#UseManualCaptcha').value==='true' && !document.querySelector('[data-manual-captcha]').hidden");
    if(!await evaluate("document.querySelector('[data-turnstile-container]').hidden"))throw new Error('No activa respaldo al bloquear script Turnstile.');

    await send('Network.setBlockedURLs',{urls:[]});
    socket.addEventListener('message',async event=>{const message=JSON.parse(event.data);if(message.method!=='Fetch.requestPaused')return;const request=message.params;try{if(request.request.url.includes('challenges.cloudflare.com/'))await send('Fetch.fulfillRequest',{requestId:request.requestId,responseCode:200,responseHeaders:[{name:'Content-Type',value:'application/javascript'}],body:Buffer.from('window.turnstile={};').toString('base64')});else{await delay(800);await send('Fetch.continueRequest',{requestId:request.requestId});}}catch{}});
    await send('Fetch.enable',{patterns:[{urlPattern:'*challenges.cloudflare.com/*',requestStage:'Request'},{urlPattern:'*/Acceso/Captcha/*',requestStage:'Request'}]});
    await navigate('/Acceso');
    await evaluate(`document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)}`);
    await evaluate('hytTurnstileFallback()');await wait("document.querySelector('#UseManualCaptcha').value==='true'");await wait(visible);await wait("document.querySelector('[data-captcha-image]').naturalWidth===260");await wait(idle);
    if(!await evaluate(`document.querySelector('#UserName').value===${JSON.stringify(fixture.user)} && document.querySelector('#Password').value===${JSON.stringify(fixture.password)}`))throw new Error('Elegir respaldo pierde credenciales escritas.');
    await navigate('/Acceso');await wait("document.querySelector('#UseManualCaptcha').value==='False'");
    await evaluate("window.qaSlowEnd=hytLoading.begin('Esperando verificación de prueba…')");
    await wait("document.querySelector('#UseManualCaptcha').value==='true'",160);
    await wait("document.querySelector('#hyt-loading')?.textContent.includes('está demorando')",160);
    await wait("document.querySelector('[data-captcha-image]').naturalWidth===260");await evaluate("qaSlowEnd()");await wait(idle);
    await send('Fetch.disable');
    await evaluate(`(()=>{document.querySelector('#UserName').value=${JSON.stringify(fixture.user)};document.querySelector('#Password').value=${JSON.stringify(fixture.password)};hytTurnstileFallback();})()`);
    if(!await evaluate(`document.querySelector('#UserName').value===${JSON.stringify(fixture.user)} && document.querySelector('#Password').value===${JSON.stringify(fixture.password)}`))throw new Error('El respaldo pierde los datos escritos.');
    await evaluate("document.querySelector('#CaptchaAnswer').value='000000';document.querySelector('form').requestSubmit()");
    await wait("document.body.textContent.includes('incorrecto o venció')");
    if(!await evaluate("document.querySelector('#Password').value==='' && !document.querySelector('[data-manual-captcha]').hidden"))throw new Error('Error captcha deja contraseña o no renueva.');
    const answer=await evaluate("fetch(document.querySelector('[data-captcha-image]').src).then(r=>r.text()).then(text=>[...new DOMParser().parseFromString(text,'image/svg+xml').querySelectorAll('text')].map(node=>node.textContent).join(''))");

    const lookupsBefore=await evaluate("fetch('/__tests/loading-count').then(r=>r.json()).then(data=>data.lookups)");
    await evaluate(`(()=>{sessionStorage.removeItem('qa-native-visible');setTimeout(()=>sessionStorage.setItem('qa-native-visible',String(!!document.querySelector('#hyt-loading')&&!document.querySelector('#hyt-loading').hidden)),350);document.querySelector('#Password').value=${JSON.stringify(fixture.password)};document.querySelector('#CaptchaAnswer').value=${JSON.stringify(answer)};const form=document.querySelector('form');form.requestSubmit();form.requestSubmit();})()`);
    await delay(400);await capture('loading-acceso');await wait("location.pathname==='/' && document.readyState==='complete'");await wait(idle);
    if(!await evaluate("sessionStorage.getItem('qa-native-visible')==='true'"))throw new Error('No mostró loading durante la respuesta de ingreso.');
    if(!await evaluate(`fetch('/__tests/loading-count').then(r=>r.json()).then(data=>data.lookups===${lookupsBefore+1})`))throw new Error('Login repetido consultó dos veces credenciales.');
    console.log('OK navegador: captcha sin claves, script Turnstile bloqueado, renovación, error, acceso y doble envío.');

    await evaluate("window.qa={};void Promise.all([fetch('/__tests/loading-delay?ms=350').then(r=>r.json()).then(()=>qa.first=true),fetch('/__tests/loading-delay?ms=1400').then(r=>r.json()).then(()=>qa.second=true)])");
    await wait(visible);await wait("qa.first && !qa.second && hytLoading.pending===1");await wait("qa.second");await wait(idle);
    await evaluate("window.qa={};void fetch('/__tests/loading-body').then(response=>{qa.headers=true;return response.json()}).then(()=>qa.body=true)");
    await wait("qa.headers && !qa.body");await wait(visible);await wait("qa.body");await wait(idle);
    for(const kind of ['error','invalid']){
        await evaluate(`window.qa={};void fetch('/__tests/loading-delay?kind=${kind}').then(r=>{if(!r.ok)throw new Error();return r.json()}).catch(()=>qa.failed=true)`);
        await wait(visible);await wait("qa.failed");await wait(idle);
    }
    await evaluate("window.qa={};window.qaAbort=new AbortController();void fetch('/__tests/loading-delay?ms=3000',{signal:qaAbort.signal}).then(r=>r.json()).catch(()=>qa.aborted=true)");
    await wait(visible);await evaluate("qaAbort.abort()");await wait("qa.aborted");await wait(idle);
    await evaluate("window.qa={};const xhr=new XMLHttpRequest();xhr.open('GET','/__tests/loading-delay?ms=900');xhr.onloadend=()=>qa.xhr=true;xhr.send()");await wait(visible);await wait("qa.xhr");await wait(idle);


    await evaluate("window.qa={};window.qaAbort=new AbortController();void fetch('/__tests/loading-body',{signal:qaAbort.signal}).then(response=>{qa.headers=true;return response.json()}).catch(()=>qa.aborted=true)");
    await wait("qa.headers");await wait(visible);await evaluate("qaAbort.abort()");await wait("qa.aborted");await wait(idle);
    await evaluate("window.qa={};window.qaXhr=new XMLHttpRequest();qaXhr.open('GET','/__tests/loading-delay?ms=3000');qaXhr.onloadend=()=>qa.xhr=true;qaXhr.send()");await wait(visible);await evaluate("qaXhr.abort()");await wait("qa.xhr");await wait(idle);
    await evaluate("(()=>{const form=document.createElement('form');form.id='qa-invalid';form.innerHTML='<input required><button type=submit>Enviar</button>';document.querySelector('main').append(form);form.requestSubmit();})()");await delay(200);
    if(!await evaluate("hytLoading.pending===0 && !document.querySelector('#qa-invalid button').disabled"))throw new Error('Formulario inválido deja espera activa.');
    const downloadFolder=path.join(profile,'downloads');await fs.mkdir(downloadFolder,{recursive:true});
    await send('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:downloadFolder});
    await evaluate("(()=>{const link=document.createElement('a');link.id='qa-download';link.textContent='Descargar prueba';link.href='/__tests/loading-delay?kind=file&ms=900&download=true';document.querySelector('main').append(link);link.click();})()");
    await wait(visible);await wait(idle);
    let downloaded=false;for(let attempt=0;attempt<30;attempt++){try{downloaded=(await fs.readFile(path.join(downloadFolder,'guia-prueba.pdf'))).subarray(0,5).toString()==='%PDF-';if(downloaded)break;}catch{}await delay(100);}if(!downloaded)throw new Error('Descarga no conserva contenido o nombre.');
    await evaluate("document.querySelector('#qa-download').href='/__tests/loading-delay?kind=error&ms=700&download=true';document.querySelector('#qa-download').click()");await wait(visible);await wait("document.querySelector('[data-hyt-download-error]')");await wait(idle);
    console.log('OK navegador: solicitudes simultáneas, cuerpo demorado, error HTTP/JSON, cancelación y XHR retiran loading.');

    await navigate('/Seguimiento/oc-cliente');await click('[data-document-open][href*="/Facturado"]');await wait("document.querySelector('#document-dialog [name=Series]')");
    const focus=await evaluate("document.activeElement?.name");
    await evaluate("window.qa={};void fetch('/__tests/loading-delay?ms=900').then(r=>r.json()).then(()=>qa.done=true)");await wait(visible);
    if(!await evaluate(`document.querySelector('#hyt-loading').matches(':popover-open') && document.activeElement?.name===${JSON.stringify(focus)} && document.querySelector('#document-dialog').open`))throw new Error('Loading no aparece sobre popup o roba foco.');
    await capture('loading-popup');await wait("qa.done");await wait(idle);

    const csrf=await evaluate("fetch('/OCCliente/1/Facturado').then(r=>r.text()).then(html=>new DOMParser().parseFromString(html,'text/html').querySelector('[name=__RequestVerificationToken]').value)");
    const postsBefore=await evaluate("fetch('/__tests/loading-count').then(r=>r.json()).then(data=>data.posts)");
    await evaluate(`(()=>{sessionStorage.removeItem('qa-post-visible');setTimeout(()=>{document.querySelector('#document-dialog [data-hyt-close]').click();sessionStorage.setItem('qa-post-visible',String(!!document.querySelector('#hyt-loading')&&!document.querySelector('#hyt-loading').hidden&&document.querySelector('#document-dialog').open));},350);const form=document.createElement('form');form.method='post';form.action='/__tests/loading-submit';for(const [name,value]of Object.entries({__RequestVerificationToken:${JSON.stringify(csrf)}})){const input=document.createElement('input');input.type='hidden';input.name=name;input.value=value;form.append(input);}const button=document.createElement('button');button.type='submit';button.name='intent';button.value='guardar';button.textContent='Guardar prueba';form.append(button);document.querySelector('#document-dialog [data-document-body]').append(form);form.requestSubmit(button);form.requestSubmit(button);})()`);
    await wait("location.pathname==='/__tests/loading-page' && document.readyState==='complete'");await wait(idle);
    if(!await evaluate(`fetch('/__tests/loading-count').then(r=>r.json()).then(data=>data.posts===${postsBefore+1} && data.intent==='guardar' && sessionStorage.getItem('qa-post-visible')==='true')`))throw new Error('Doble envío nativo o submitter perdido.');
    await evaluate("(()=>{sessionStorage.removeItem('qa-get-locked');setTimeout(()=>sessionStorage.setItem('qa-get-locked',String(!!document.querySelector('#qa-search button')?.disabled&&!!document.querySelector('#hyt-loading')&&!document.querySelector('#hyt-loading').hidden)),350);const form=document.createElement('form');form.id='qa-search';form.method='get';form.action='/__tests/loading-page';form.innerHTML='<input name=ms value=900><button type=submit>Buscar</button>';document.querySelector('main').append(form);form.requestSubmit();})()");
    await wait("location.search.includes('ms=900') && document.readyState==='complete'");await wait(idle);
    if(!await evaluate("sessionStorage.getItem('qa-get-locked')==='true'"))throw new Error('Consulta GET sin loading o sin bloqueo de doble envío.');
    await send('Page.navigateToHistoryEntry',{entryId:(await send('Page.getNavigationHistory')).entries.at(-2).id});
    await wait("document.readyState==='complete' && !document.querySelector('#qa-search button')?.disabled");await wait(idle);
    console.log('OK navegador: envíos nativos, submitter, consultas GET y retorno con controles habilitados.');
    if(errors.length)throw new Error('Errores JavaScript: '+errors.join(';'));
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
