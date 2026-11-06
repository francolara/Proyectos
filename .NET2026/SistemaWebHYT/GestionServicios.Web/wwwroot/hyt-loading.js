// Firma: FRANCO LARA - 05/11/2026 | Espera común para navegación, formularios, AJAX y descargas; concurrencia, errores y retorno del navegador.
(() => {
    if (globalThis.hytLoading) return;
    const tasks = new Map(), forms = new Map(), submitting = new Set(), downloads = new Set();
    let notice, label, showTimer, slowTimer, navigation;
    function ensure() {
        if (notice) return;
        notice = document.createElement('div');
        notice.id = 'hyt-loading'; notice.className = 'hyt-loading';
        notice.setAttribute('popover', 'manual'); notice.setAttribute('role', 'status');
        notice.setAttribute('aria-live', 'polite'); notice.setAttribute('aria-atomic', 'true'); notice.hidden = true;
        const spinner = document.createElement('span'); spinner.className = 'hyt-loading-spinner'; spinner.setAttribute('aria-hidden', 'true');
        label = document.createElement('strong'); notice.append(spinner, label); document.body.append(notice);
    }
    function update() {
        if (!tasks.size) {
            clearTimeout(showTimer); clearTimeout(slowTimer); showTimer = slowTimer = null;
            if (notice) { if (notice.hidePopover && notice.matches(':popover-open')) notice.hidePopover(); notice.hidden = true; }
            document.querySelector('main')?.removeAttribute('aria-busy'); return;
        }
        document.querySelector('main')?.setAttribute('aria-busy', 'true');
        if (label) label.textContent = [...tasks.values()].at(-1);
        if (!showTimer) showTimer = setTimeout(() => {
            ensure(); label.textContent = [...tasks.values()].at(-1) || 'Cargando…'; notice.hidden = false;
            if (notice.showPopover) notice.showPopover();
        }, 150);
        if (!slowTimer) slowTimer = setTimeout(() => { if (label && tasks.size) label.textContent = 'La respuesta está demorando. Seguimos esperando…'; }, 12000);
    }
    function begin(message = 'Cargando…') {
        const id = Symbol(); tasks.set(id, message); update(); let ended = false;
        return () => { if (!ended) { ended = true; tasks.delete(id); update(); } };
    }
    function restoreForms() {
        for (const [form, buttons] of forms) {
            form.removeAttribute('aria-busy'); delete form.dataset.hytSubmitting;
            for (const button of buttons) { button.disabled = false; button.classList.remove('hyt-is-submitting'); button.removeAttribute('aria-disabled'); }
        }
        forms.clear();
    }
    function reset() { tasks.clear(); navigation = null; submitting.clear(); restoreForms(); update(); }
    function navigate(message = 'Cargando página…') { if (!navigation) navigation = begin(message); }
    globalThis.hytLoading = { begin, navigation: navigate, get pending() { return tasks.size; } };

    const nativeFetch = globalThis.fetch.bind(globalThis);
    globalThis.fetch = async function(input, options) {
        let url;
        try { url = new URL(input instanceof Request ? input.url : input, location.href); } catch { return nativeFetch(input, options); }
        if (url.origin !== location.origin) return nativeFetch(input, options);
        const method = (options?.method || (input instanceof Request ? input.method : 'GET')).toUpperCase();
        const end = begin(method === 'GET' || method === 'HEAD' ? 'Consultando…' : 'Procesando solicitud…');
        try {
            const response = await nativeFetch(input, options); let reading = false;
            for (const name of ['json', 'text', 'blob', 'arrayBuffer', 'formData', 'bytes']) {
                if (typeof response[name] !== 'function') continue;
                const read = response[name].bind(response);
                response[name] = async (...args) => { reading = true; try { return await read(...args); } finally { end(); } };
            }
            setTimeout(() => { if (!reading) end(); }, 0); // Respuestas sin cuerpo o rechazadas por el consumidor.
            return response;
        } catch (error) { end(); throw error; }
    };
    const send = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function(...args) {
        const end = begin('Procesando solicitud…'); this.addEventListener('loadend', end, { once: true });
        try { return send.apply(this, args); } catch (error) { end(); throw error; }
    };
    document.addEventListener('submit', event => {
        const form = event.target;
        if (!(form instanceof HTMLFormElement)) return;
        if (forms.has(form) || submitting.has(form)) { event.preventDefault(); return; }
        submitting.add(form);
        setTimeout(() => {
            submitting.delete(form);
            const target=event.submitter?.formTarget||form.target;
            if (event.defaultPrevented || !form.checkValidity() || form.method === 'dialog' || target && target !== '_self' || form.hasAttribute('data-hyt-no-loading')) return;
            const buttons = [...form.querySelectorAll('button:not([type=button]):not([type=reset]),input[type=submit]')].filter(button => !button.disabled);
            forms.set(form, buttons); form.dataset.hytSubmitting = 'true'; form.setAttribute('aria-busy', 'true');
            // No deshabilitar un submitter con nombre: su valor debe conservarse en el envío.
            buttons.forEach(button => { if (!button.name) button.disabled = true; button.setAttribute('aria-disabled', 'true'); button.classList.add('hyt-is-submitting'); });
            navigate(form.method.toLowerCase() === 'get' ? 'Consultando…' : form.action.includes('/Acceso') ? 'Verificando acceso…' : 'Guardando…');
        });
    }, true);
    document.addEventListener('click', event => {
        const link = event.target.closest?.('a[href]');
        if (!link || event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        const requested = new URL(link.href, location.href);
        if (!event.defaultPrevented && requested.origin === location.origin && requested.searchParams.get('download')?.toLowerCase() === 'true' && !link.target) {
            event.preventDefault();
            if (downloads.has(link.href)) return;
            downloads.add(link.href);
            const end = begin('Descargando archivo…');
            (async () => {
                try {
                    const response = await fetch(link.href);
                    if (!response.ok || response.redirected) throw new Error();
                    const blob = await response.blob(), disposition = response.headers.get('content-disposition') || '';
                    let name = disposition.match(/filename="([^"\r\n]+)"/i)?.[1] || 'documento';
                    const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
                    if (encoded) { try { name = decodeURIComponent(encoded); } catch {} }
                    const object = URL.createObjectURL(blob), anchor = document.createElement('a');
                    anchor.href = object; anchor.download = name; document.body.append(anchor); anchor.click(); anchor.remove();
                    setTimeout(() => URL.revokeObjectURL(object), 60000);
                    link.parentElement?.querySelector('[data-hyt-download-error]')?.remove();
                } catch {
                    let message = link.parentElement?.querySelector('[data-hyt-download-error]');
                    if (!message) { message = document.createElement('p'); message.dataset.hytDownloadError = ''; message.setAttribute('role', 'alert'); link.after(message); }
                    message.textContent = 'No se pudo descargar el archivo. Inténtalo nuevamente.';
                } finally { downloads.delete(link.href); end(); }
            })();
            return;
        }
        setTimeout(() => {
            if (event.defaultPrevented || link.hasAttribute('download') || link.hasAttribute('data-hyt-no-loading') || link.target && link.target !== '_self') return;
            const url = new URL(link.href, location.href);
            if (url.origin !== location.origin || !['https:', 'http:'].includes(url.protocol) || url.searchParams.get('download')?.toLowerCase() === 'true') return;
            if (url.pathname === location.pathname && url.search === location.search && url.hash) return;
            navigate();
        });
    });
    window.addEventListener('pageshow', event => { if (event.persisted) reset(); else restoreForms(); }); window.addEventListener('pagehide', reset);
    window.addEventListener('keydown', event => { if (event.key === 'Escape' && navigation) { navigation(); navigation = null; restoreForms(); } });
})();
