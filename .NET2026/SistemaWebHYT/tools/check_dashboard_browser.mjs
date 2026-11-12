// Flujos MVC reales con repositorios de prueba aislados; nunca conecta SQL.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {setTimeout as delay} from 'node:timers/promises';
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
    await wait("document.querySelector('[data-chart=trend] path') && document.querySelector('[data-chart=states] rect') && document.querySelector('[data-chart=donut] circle')");
    if (!await evaluate("document.querySelectorAll('[data-chart=trend] path[stroke]').length===2 && document.querySelector('[data-chart=trend] path[fill-opacity]') && document.querySelectorAll('[data-sparkline] path').length===4")) throw new Error('Faltan las dos curvas, el área o los minigráficos.');
    const animations = await evaluate("['.hyt-dashboard-metrics article','.hyt-chart-reveal','.hyt-chart-bar','.hyt-dashboard-donut-segments'].map(selector=>getComputedStyle(document.querySelector(selector)).animationName)");
    if (animations.some(name => !name.startsWith('hyt-dashboard-'))) throw new Error('Falta una animación de entrada.');
    for (const key of ['guide', 'order', 'both', 'acceptance']) {
        await click('[data-coverage-select="' + key + '"]');
        const state = await evaluate("(() => { const panel=document.querySelector('[data-coverage-view]:not([hidden])');return {pressed:document.querySelector('[data-coverage-select][aria-pressed=true]').dataset.coverageSelect,view:panel.dataset.coverageView,animation:getComputedStyle(panel.querySelector('.hyt-budget-coverage-fill')).animationName};})()");
        if (state.pressed !== key || state.view !== key || state.animation !== 'hyt-budget-coverage-grow') throw new Error('La cobertura no cambia o no anima: ' + key);
    }
    await click('[data-chart-series="Presupuestos"]');if(!await evaluate("document.querySelectorAll('[data-chart=trend] path[stroke]').length===1"))throw new Error('No oculta serie.');
    await click('[data-chart-series="Ordenes"]');if(!await evaluate("document.querySelector('[data-chart-note]').textContent.includes('Selecciona una serie')"))throw new Error('Estado sin series no identificado.');
    await click('[data-chart-series="Presupuestos"]');await click('[data-chart-series="Ordenes"]');
    await evaluate("const select=document.querySelector('select[aria-label=\"Consultar intervalo del gráfico\"]');select.value=String(select.options.length-1);select.dispatchEvent(new Event('change',{bubbles:true}))");
    if(!await evaluate("document.querySelector('[data-chart-readout]').textContent.includes('Presupuestos: 5') && document.querySelector('[data-chart-readout]').textContent.includes('OC del cliente: 2')"))throw new Error('Detalle del último intervalo no coincide.');
    await click('[data-chart-series="Confirmaciones"]');await click('[data-chart-series="Marcados"]');
    if(!await evaluate("document.querySelectorAll('[data-chart=trend] path[stroke]').length===4 && [...document.querySelectorAll('[data-chart=trend] path[stroke]')].filter(path=>path.getAttribute('stroke-dasharray')!=='none').length===2"))throw new Error('Faltan series de eventos diferenciadas.');
    for(const group of ['week','month']){await evaluate(`document.querySelector('[data-chart-group]').value=${JSON.stringify(group)};document.querySelector('[data-chart-group]').dispatchEvent(new Event('change',{bubbles:true}))`);if(!await evaluate("['Confirmación del cliente','Facturado'].every(label=>[...document.querySelectorAll('[data-chart=trend] circle')].filter(point=>point.getAttribute('aria-label').includes(label)).reduce((sum,point)=>sum+Number(point.getAttribute('aria-label').split(': ').at(-1)),0)===7)"))throw new Error('Agrupación pierde movimientos '+group);}
    await evaluate("document.querySelector('[data-chart-group]').value='day';document.querySelector('[data-chart-group]').dispatchEvent(new Event('change',{bubbles:true}));document.querySelector('[data-chart=trend] circle').dispatchEvent(new MouseEvent('click',{bubbles:true}))");
    if(!await evaluate("!document.querySelector('[data-chart-link]').hidden && document.querySelector('[data-chart-link]').getAttribute('href').includes('/Presupuestos?from=') && document.body.textContent.includes('Expedientes por completar') && document.querySelector('a[href*=queuePage]')"))throw new Error('Navegación por intervalo o bandeja incompleta.');
    for(const [label,width,height] of [['escritorio',1440,1100],['tableta',870,950],['movil',390,844],['movil-estrecho',320,760]]){
        await send('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await evaluate('window.scrollTo(0,0);new Promise(requestAnimationFrame)');
        if(await evaluate('document.documentElement.scrollWidth>innerWidth+1'))throw new Error('Dashboard desborda '+label);
        if (process.env.HYT_DASHBOARD_PREVIEW_DIR) {
            const screenshot = await send('Page.captureScreenshot', {format:'png'});
            await fs.writeFile(path.join(process.env.HYT_DASHBOARD_PREVIEW_DIR, 'dashboard-' + label + '.png'), Buffer.from(screenshot.data, 'base64'));
        }
    }
    await navigate('/?from=2026-01-01&to=2026-12-31');
    if(!await evaluate("document.querySelector('select[aria-label=\"Consultar intervalo del gráfico\"]').options.length<=90"))throw new Error('Intervalos sin acotar.');
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate('/?client=Cliente&store=Centro');
    await wait("document.querySelector('[data-chart=states] rect')");
    const hasContext = await evaluate("document.querySelector('.hyt-dashboard-filters [name=client]').value==='Cliente' && document.querySelector('.hyt-dashboard-filters [name=store]').value==='Centro' && document.querySelector('a[href*=queuePage]').href.includes('client=Cliente')");
    if (!hasContext) throw new Error('Los filtros de Inicio o la bandeja perdieron el contexto comercial.');
    await wait("document.querySelector('.hyt-dashboard-charts').getAnimations({subtree:true}).every(animation => animation.playState === 'finished')");
    await click('[data-chart=states] rect');
    await wait("location.pathname==='/Presupuestos'");
    if (!await evaluate("new URL(location.href).searchParams.get('state')==='BORRADOR' && new URL(location.href).searchParams.get('client')==='Cliente' && new URL(location.href).searchParams.get('store')==='Centro'")) throw new Error('La barra no conserva estado y filtros.');
    await navigate('/');
    const themedCharts = await evaluate("(() => {const root=document.documentElement; const panel=document.querySelector('.hyt-dashboard');const initial=root.dataset.hytTheme; root.dataset.hytTheme='light'; const light=getComputedStyle(panel).getPropertyValue('--hyt-dashboard-blue'); root.dataset.hytTheme='dark'; const dark=getComputedStyle(panel).getPropertyValue('--hyt-dashboard-blue'); if(initial)root.dataset.hytTheme=initial;else delete root.dataset.hytTheme; return light.trim()!=='' && dark.trim()!=='' && light!==dark;})()");
    if (!themedCharts) throw new Error('La paleta del gráfico no se adapta al tema.');
    await send('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
    await navigate('/');
    const reduced = await evaluate("['.hyt-dashboard-metrics article','.hyt-chart-reveal','.hyt-chart-bar','.hyt-dashboard-donut-segments','.hyt-budget-coverage-fill'].every(selector=>getComputedStyle(document.querySelector(selector)).animationName==='none')");
    if (!reduced) throw new Error('El dashboard no respeta movimiento reducido.');
    await click('a[href*="pendingInvoice=true"]');
    await wait("location.pathname==='/Reportes'");
    if (!await evaluate("new URL(location.href).searchParams.get('kind')==='OC_CLIENTE' && new URL(location.href).searchParams.get('pendingInvoice')==='true'")) throw new Error('El acceso Por facturar no conserva el filtro.');
    await navigate('/?client=HYT-QA-SIN-DATOS');
    await wait("document.querySelector('[data-chart-note]').textContent.includes('No hay registros')");
    if (!await evaluate("document.querySelector('[data-coverage-view=acceptance]').textContent.includes('No hay presupuestos') && document.querySelector('.hyt-dashboard-queue tbody').textContent.includes('Sin expedientes') && !document.querySelector('[data-sparkline]:not([hidden])')")) throw new Error('El estado vacío del dashboard es incompleto.');
    console.log('OK dashboard MVC: tarjetas, minigráficos, curvas, filtros, expediente, cobertura, animaciones y movimiento reducido.');
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
