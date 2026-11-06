// Firma: FRANCO LARA - 05/11/2026 | Precios en ventana contextual, mensajes del mismo origen y actualización parcial sin perder el material editado.
import {setDialog} from './hyt-ui.js';
const region=document.querySelector('[data-material-tariffs]'),dialog=document.querySelector('#material-price-dialog');
if(region&&dialog){
    const frame=dialog.querySelector('iframe'),list=region.querySelector('[data-material-price-list]'),status=region.querySelector('[data-material-price-status]'),material=region.dataset.materialTariffs;
    let busy=false,endLoading,request=0;
    const page=()=>Number(list.querySelector('[data-price-page]')?.dataset.pricePage||1);
    async function refresh(selected=page()){
        const current=++request;status.textContent='Consultando precios…';
        try{const response=await fetch(`/Maestros/materiales/${encodeURIComponent(material)}/Tarifas?page=${selected}`,{headers:{'X-Requested-With':'XMLHttpRequest'}});if(!response.ok||response.redirected)throw new Error();const html=await response.text();if(current!==request)return;list.innerHTML=html;status.textContent='Precios actualizados.';}
        catch{if(current===request)status.textContent='No se pudieron actualizar los precios. Vuelve a consultar la ficha.';}
    }
    function close(){if(busy)return;endLoading?.();endLoading=null;setDialog(dialog,false);frame.src='about:blank';}
    document.addEventListener('click',event=>{
        const pager=event.target.closest('[data-material-price-page]');if(pager&&region.contains(pager)){event.preventDefault();refresh(Number(pager.dataset.materialPricePage));return;}
        const link=event.target.closest('[data-material-price-open]');if(!link||!region.contains(link))return;
        event.preventDefault();if(dialog.open)return;
        const url=new URL(link.href,location.href);if(url.origin!==location.origin)return;
        busy=true;setDialog(dialog,true,true,close);endLoading=globalThis.hytLoading?.begin('Cargando tarifario…');frame.src=url.href;
    });
    frame.addEventListener('load',()=>{endLoading?.();endLoading=null;if(dialog.open){const heading=frame.contentDocument?.querySelector('h1');if(heading)dialog.querySelector('#material-price-title').textContent=heading.textContent;busy=false;setDialog(dialog,true,false,close);}});
    dialog.querySelector('[data-material-price-close]').addEventListener('click',close);
    window.addEventListener('message',event=>{
        if(event.origin!==location.origin||event.source!==frame.contentWindow||String(event.data?.materialId)!==material)return;
        if(event.data.type==='hyt-price-busy'){busy=!!event.data.busy;if(dialog.open)setDialog(dialog,true,busy,close);}
        if(event.data.type==='hyt-price-cancel'){if(!busy){close();refresh();}}
        if(event.data.type==='hyt-price-saved'){busy=false;close();refresh(event.data.resetPage?1:page());}
    });
}
