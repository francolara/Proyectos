// Inspección de la pantalla de acceso, sin iniciar sesión ni conectar SQL.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {setTimeout as delay} from 'node:timers/promises';
import {fileURLToPath} from 'node:url';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const targetUrl = process.argv[2] || 'http://127.0.0.1:5187/Acceso';
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
    const loaded = new Promise(resolve => events.set('Page.loadEventFired', resolve));
    await send('Page.navigate', {url: targetUrl}); await loaded;
    // Verificar el módulo real del popup en Chrome, sin cuentas ni datos de negocio.
    await send('Runtime.evaluate', {awaitPromise: true, expression: `(async()=>{
        const module=await import('/hyt-ui.js');
        const trigger=document.createElement('button'); trigger.textContent='Prueba de foco'; document.body.append(trigger); trigger.focus();
        const dialog=document.createElement('dialog'); dialog.innerHTML='<button autofocus>Primero</button><button>Segundo</button>'; document.body.append(dialog);
        window.__hytDialogCheck={module,trigger,dialog,calls:0};
        const reference=async()=>{window.__hytDialogCheck.calls++;module.setDialog(dialog,false,false,reference);};
        window.__hytDialogCheck.reference=reference; module.setDialog(dialog,true,false,reference);
    })()`});
    for (let i = 0; i < 5; i++) {
        await send('Input.dispatchKeyEvent', {type: 'keyDown', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9});
        await send('Input.dispatchKeyEvent', {type: 'keyUp', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9});
        const focus = await send('Runtime.evaluate', {returnByValue: true, expression: '__hytDialogCheck.dialog.contains(document.activeElement) || document.activeElement === document.body'});
        if (!focus.result.value) throw new Error('El foco salió del popup modal.');
    }
    await send('Input.dispatchKeyEvent', {type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27});
    await send('Input.dispatchKeyEvent', {type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27});
    const restored = await send('Runtime.evaluate', {returnByValue: true, expression: '!__hytDialogCheck.dialog.open && __hytDialogCheck.calls===1 && document.activeElement===__hytDialogCheck.trigger'});
    if (!restored.result.value) throw new Error('Escape no cerró el popup o no devolvió el foco.');
    await send('Runtime.evaluate', {expression: '__hytDialogCheck.module.setDialog(__hytDialogCheck.dialog,true,true,__hytDialogCheck.reference)'});
    await send('Input.dispatchKeyEvent', {type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27});
    await send('Input.dispatchKeyEvent', {type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27});
    const busy = await send('Runtime.evaluate', {returnByValue: true, expression: '({open:__hytDialogCheck.dialog.open,calls:__hytDialogCheck.calls})'});
    if (!busy.result.value.open || busy.result.value.calls !== 1) throw new Error('Escape cerró un popup ocupado: ' + JSON.stringify(busy.result.value));
    await send('Runtime.evaluate', {expression: '__hytDialogCheck.module.setDialog(__hytDialogCheck.dialog,false,false,__hytDialogCheck.reference);__hytDialogCheck.dialog.remove();__hytDialogCheck.trigger.remove();delete window.__hytDialogCheck;'});
    console.log('OK navegador: popup nativo, Escape, retorno de foco y bloqueo durante guardado.');
    const calendar = await send('Runtime.evaluate', {awaitPromise: true, returnByValue: true, expression: `(async()=>{
        await Promise.all([
            new Promise(resolve=>{const link=document.createElement('link');link.rel='stylesheet';link.href='/hyt-components.css';link.onload=resolve;document.head.append(link);}),
            new Promise(resolve=>{const script=document.createElement('script');script.src='/hyt-dates.js';script.onload=resolve;document.head.append(script);})
        ]);
        const form=document.createElement('form');form.id='hyt-date-test';form.innerHTML='<label for="hyt-check-date">Fecha</label><input id="hyt-check-date" type="date" name="Fecha" value="2026-11-03">';document.body.append(form);
        window.hytInitDates(form);form.querySelector('.app-date-picker-trigger').click();
        const panel=document.querySelector('.app-date-picker-panel');
        const monday=panel.querySelector('.app-date-picker-weekdays span').textContent==='LU';
        panel.querySelector('[data-value="2026-11-12"]').click();
        return monday&&form.querySelector('.app-date-picker-display').value==='12/11/2026'&&new FormData(form).get('Fecha')==='2026-11-12';
    })()`});
    if (!calendar.result.value) throw new Error('El calendario MVC no conserva formato y valor del formulario.');
    await send('Runtime.evaluate', {expression: "document.querySelector('#hyt-date-test .app-date-picker-trigger').click()"});
    await send('Input.dispatchKeyEvent', {type:'keyDown',key:'Escape',code:'Escape',windowsVirtualKeyCode:27});
    await send('Input.dispatchKeyEvent', {type:'keyUp',key:'Escape',code:'Escape',windowsVirtualKeyCode:27});
    const calendarClosed = await send('Runtime.evaluate', {returnByValue:true,expression:"document.querySelector('.app-date-picker-panel').hidden && document.activeElement===document.querySelector('#hyt-date-test .app-date-picker-trigger')"});
    if (!calendarClosed.result.value) throw new Error('El calendario MVC no devuelve el foco al cerrarse con Escape.');
    await send('Runtime.evaluate', {expression:"document.getElementById('hyt-date-test').remove();document.querySelectorAll('.app-date-picker-panel').forEach(p=>p.remove());"});
    console.log('OK navegador: calendario MVC dd/MM/yyyy, valor ISO, lunes y retorno de foco.');

    for (const [label, width, height] of [['escritorio', 1440, 1000], ['movil', 390, 844]]) {
        await send('Emulation.setDeviceMetricsOverride', {width, height, deviceScaleFactor: 1, mobile: false});
        await send('Runtime.evaluate', {expression: 'new Promise(requestAnimationFrame)', awaitPromise: true});
        const result = await send('Runtime.evaluate', {returnByValue: true, expression: `({width:innerWidth, scroll:document.documentElement.scrollWidth, title:document.title, form:!!document.querySelector('form[method="post"]'), fields:[...document.querySelectorAll('.hyt-field input')].map(e=>({left:e.getBoundingClientRect().left,right:e.getBoundingClientRect().right}))})`});
        const state = result.result.value;
        if (state.scroll > width + 1 || !state.form || state.fields.some(field => field.left < 0 || field.right > width + 1)) throw new Error(`Desbordamiento o formulario ausente: ${JSON.stringify(state)}`);
        const toggle = await send('Runtime.evaluate', {returnByValue: true, expression: `(()=>{const button=document.querySelector('[data-hyt-password-toggle]');if(!button)return false;button.click();const shown=document.getElementById('Password').type==='text';button.click();return shown&&document.getElementById('Password').type==='password';})()`});
        if (!toggle.result.value) throw new Error('El control de mostrar/ocultar contraseña no funciona.');
        const screenshot = await send('Page.captureScreenshot', {format: 'png'});
        const output = path.join(project, 'entregables', 'preview', `acceso-${label}.png`);
        await fs.mkdir(path.dirname(output), {recursive: true}); await fs.writeFile(output, Buffer.from(screenshot.data, 'base64'));
        console.log(`OK navegador: acceso ${label} ${width}x${height}, sin desbordamiento.`);
    }
    await send('Browser.close');
} finally {
    socket?.close();
    if (chrome.exitCode === null) chrome.kill();
    const resolved = path.resolve(profile);
    if (!resolved.startsWith(tempRoot + path.sep) || !path.basename(resolved).startsWith('hyt-browser-')) throw new Error('Ruta temporal inesperada.');
    // Esperar cierre de Chrome antes de quitar exclusivamente su perfil temporal.
    for (let i = 0; i < 20; i++) { try { await fs.rm(resolved, {recursive: true, force: true}); break; } catch { await delay(100); } }
}
