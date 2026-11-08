// Firma: FRANCO LARA - 06/11/2026 | Respaldo manual automático solo por error o falta de carga; respeta Turnstile interactivo y su recuperación.
// Firma: FRANCO LARA - 07/11/2026 | La visibilidad de contraseña se gestiona en el control compartido.
const errors = document.querySelector(".validation-summary-errors");
if (errors) errors.focus();

// Firma: FRANCO LARA - 05/11/2026 | Activa y renueva el captcha ante error, ausencia o demora de Turnstile; conserva datos y espera la imagen.
(() => {
    const box=document.querySelector('[data-manual-captcha]');if(!box)return;
    const mode=document.querySelector('#UseManualCaptcha'),answer=document.querySelector('#CaptchaAnswer'),id=document.querySelector('#CaptchaId'),image=box.querySelector('[data-captcha-image]'),status=box.querySelector('[data-captcha-status]'),refresh=box.querySelector('[data-captcha-refresh]');
    let imageEnd,startupTimer;
    let healthy=false;
    function watchImage(){if(box.hidden||image.complete)return;if(!imageEnd)imageEnd=globalThis.hytLoading?.begin('Cargando código…');}
    function imageFinished(){imageEnd?.();imageEnd=null;}
    image.addEventListener('load',imageFinished);
    function activate(focus=false){
        mode.value='true';box.hidden=false;answer.disabled=false;answer.required=true;
        const widget=document.querySelector('[data-turnstile-container]');if(widget)widget.hidden=true;
        image.loading='eager';watchImage();if(focus)answer.focus();
    }
    function automatic(){
        healthy=true;clearTimeout(startupTimer);mode.value='false';box.hidden=true;answer.disabled=true;answer.required=false;imageFinished();
        const widget=document.querySelector('[data-turnstile-container]');if(widget)widget.hidden=false;
    }
    globalThis.hytTurnstileReady=automatic;
    globalThis.hytTurnstileSuccess=automatic;
    globalThis.hytTurnstileFallback=()=>{healthy=false;clearTimeout(startupTimer);activate();return true;};
    document.querySelector('#hyt-turnstile-script')?.addEventListener('error',()=>activate());
    if(mode.value.toLowerCase()!=='true')startupTimer=setTimeout(()=>{if(!healthy&&!document.querySelector('[name=cf-turnstile-response]')?.value)activate();},12000);
    let refreshing=false;
    refresh.addEventListener('click',async()=>{
        if(refreshing)return;refreshing=true;refresh.disabled=true;status.textContent='Generando otro código…';
        const end=globalThis.hytLoading?.begin('Cargando código…'),submit=box.closest('form').querySelector('button[type=submit]'),wasDisabled=submit.disabled;submit.disabled=true;answer.disabled=true;
        try{
            const response=await fetch('/Acceso/Captcha/Nuevo',{headers:{Accept:'application/json'}});
            if(!response.ok||response.redirected)throw new Error();
            const data=await response.json(),url=new URL(data.imageUrl,location.origin);if(url.origin!==location.origin)throw new Error();
            id.value=data.id;answer.value='';image.loading='eager';image.src=url.href;await image.decode();status.textContent='Código renovado. Vence en cinco minutos.';answer.disabled=false;answer.focus();
        }catch{status.textContent='No se pudo renovar el código. Inténtalo nuevamente.';}
        finally{end?.();refreshing=false;refresh.disabled=false;submit.disabled=wasDisabled;answer.disabled=mode.value.toLowerCase()!=='true';}
    });
    watchImage();
    image.addEventListener('error',()=>{imageFinished();status.textContent='La imagen no está disponible o venció. Pulsa Otro código.';});
})();
