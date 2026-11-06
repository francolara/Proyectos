// Firma: FRANCO LARA - 05/11/2026 | Notifica guardado/cierre al material del mismo origen y protege la espera del formulario.
(()=>{
    if(parent===window)return;
    const materialId=document.body.dataset.priceFrameMaterial;
    const send=(type,extra={})=>parent.postMessage({type,materialId,...extra},location.origin);
    if(document.body.hasAttribute('data-price-frame-saved'))send('hyt-price-saved',{resetPage:document.body.dataset.priceResetPage.toLowerCase()==='true'});
    document.querySelectorAll('[data-price-frame-cancel]').forEach(button=>button.addEventListener('click',()=>{if(!document.querySelector('[data-hyt-submitting=true]'))send('hyt-price-cancel');}));
    document.addEventListener('submit',event=>setTimeout(()=>{if(!event.defaultPrevented&&event.target.checkValidity())send('hyt-price-busy',{busy:true});},0));
    window.addEventListener('pageshow',()=>send('hyt-price-busy',{busy:false}));
    window.addEventListener('keydown',event=>{if(event.key==='Escape'&&!document.querySelector('[data-hyt-submitting=true]'))send('hyt-price-busy',{busy:false});});
})();
