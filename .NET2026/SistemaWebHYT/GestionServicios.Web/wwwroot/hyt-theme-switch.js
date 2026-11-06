// Firma: FRANCO LARA - 05/11/2026 | Tema persistente, preferencia del sistema y sincronización de ventanas del mismo origen.
(() => {
    const key = 'hyt.theme';
    const media = window.matchMedia('(prefers-color-scheme: dark)');
    let choice;
    try { const saved = localStorage.getItem(key); if (saved === 'dark' || saved === 'light') choice = saved; } catch { }
    function apply(theme) {
        document.documentElement.dataset.hytTheme = theme;
        document.querySelectorAll('[data-hyt-theme-toggle]').forEach(button => {
            button.setAttribute('aria-pressed', String(theme === 'dark'));
            button.querySelector('[data-hyt-theme-label]').textContent = theme === 'dark' ? 'Modo claro' : 'Modo oscuro';
            button.title = theme === 'dark' ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro';
        });
        document.querySelectorAll('iframe').forEach(frame => {
            try { if (frame.contentWindow?.location.origin === location.origin) frame.contentWindow.postMessage({type:'hyt:theme',theme}, location.origin); } catch { }
        });
    }
    function current() { return choice ?? (media.matches ? 'dark' : 'light'); }
    apply(current());
    document.addEventListener('DOMContentLoaded', () => {
        apply(current());
        document.querySelectorAll('[data-hyt-theme-toggle]').forEach(button => button.addEventListener('click', () => {
            choice = document.documentElement.dataset.hytTheme === 'dark' ? 'light' : 'dark';
            try { localStorage.setItem(key, choice); } catch { }
            apply(choice);
        }));
        document.querySelectorAll('iframe').forEach(frame => frame.addEventListener('load', () => apply(current())));
    });
    media.addEventListener('change', () => { if (!choice) apply(current()); });
    window.addEventListener('storage', event => {
        if (event.key !== key && event.key !== null) return;
        choice = event.newValue === 'dark' || event.newValue === 'light' ? event.newValue : undefined;
        apply(current());
    });
    window.addEventListener('message', event => {
        if (window.parent === window || event.source !== window.parent || event.origin !== location.origin || event.data?.type !== 'hyt:theme' || !['dark','light'].includes(event.data.theme)) return;
        choice = event.data.theme; apply(choice);
    });
})();
