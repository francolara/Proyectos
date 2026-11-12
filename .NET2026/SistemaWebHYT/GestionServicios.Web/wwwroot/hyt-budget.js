// Firma: FRANCO LARA - 10/11/2026 | Precarga la tarifa del material y el código SAP del cliente.
// Firma: FRANCO LARA - 06/11/2026 | Tarifa del material disponible sin cliente; iconos compactos con ayudas accesibles. Selección del nombre del tipo de servicio y títulos SAP; sin aviso al completar cantidades ni confirmación de moneda.
// Firma: FRANCO LARA - 03/11/2026 | Presupuestos MVC: ayudas, partidas, previsualización y revalorización explícita.
// Firma: FRANCO LARA - 05/11/2026 | Direcciones al seleccionar cliente/tienda y ayudas de zona/especialidad. Búsquedas por fila, unidad del material, tarifa filtrada y conversión neta con edición del precio.
import {setDialog} from './hyt-ui.js';

const form=document.querySelector('[data-budget-form]');
if(form){
    const editable=form.dataset.readOnly!=='true';
    const body=form.querySelector('[data-budget-lines]');
    const dialog=document.querySelector('#budget-lookup');
    const feedback=form.querySelector('[data-default-status]');
    const calculation=form.querySelector('[data-calculation-status]');
    const query=dialog.querySelector('[data-budget-search]');
    const lookupBody=dialog.querySelector('[data-budget-lookup-body]');
    const lookupStatus=dialog.querySelector('[data-budget-lookup-status]');
    const previous=dialog.querySelector('[data-budget-prev]'),next=dialog.querySelector('[data-budget-next]');
    const prices=new WeakMap();
    let lookupMode='',reference=null,rowTarget=null,page=1,busy=false,lookupRequest=0,calculationRequest=0,timer,metadataRequest=0,sapPending=false;
    const header=name=>form.querySelector(`[name="Draft.Revision[${name}]"]`);
    const field=(row,name)=>row.querySelector(`[name$=".${name}"]`);
    const number=value=>value===null||value===undefined?'Pendiente':new Intl.NumberFormat('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2}).format(value);
    function renumber(){
        body.querySelectorAll('[data-budget-line]').forEach((row,index)=>{
            row.querySelector('[data-line-number]').textContent=String(index+1);
            row.querySelectorAll('[name]').forEach(input=>input.name=input.name.replace(/Draft\.Partidas\[(?:\d+|__index__)\]/,`Draft.Partidas[${index}]`));
            row.querySelectorAll('[id]').forEach(input=>input.id=input.id.replace(/budget-line-(?:\d+|__index__)-/,`budget-line-${index}-`));
            row.querySelectorAll('[for]').forEach(label=>label.htmlFor=label.htmlFor.replace(/budget-line-(?:\d+|__index__)-/,`budget-line-${index}-`));
        });
        form.querySelector('[data-empty-lines]').hidden=body.children.length>0;
    }
    function sync(row){
        field(row,'OrigenPrecio').value='MANUAL';field(row,'CodigoTipoIGVOrigen').value='SIN_IGV';field(row,'PrecioManual').readOnly=false;
        field(row,'MotivoPrecioManual').value='Precio neto del detalle del presupuesto';
        if(!prices.has(row))prices.set(row,new Map([[field(row,'CodigoMonedaOrigen').value,field(row,'PrecioManual').value]]));
    }
    function netPrice(item){const value=item.price;if(value===null||value===undefined)return '';const rate=Number(form.dataset.taxRate);if(item.taxTreatment==='CON_IGV'&&(!form.dataset.taxRate||!Number.isFinite(rate)))throw new Error('Selecciona una fecha con IGV vigente.');return (Math.round((Number(value)/(item.taxTreatment==='CON_IGV'?1+rate:1)+Number.EPSILON)*100)/100).toFixed(2);}
    function applyTariff(row,item){
        for(const [name,value]of Object.entries({OrigenPrecio:'MANUAL',PartidaTarifarioId:item.tariffLineId,MaterialId:item.materialId,CodigoInterno:item.internalCode,Posicion:item.position,CodigoMaterialSAP:item.sapCode,Descripcion:item.description,Unidad:item.unit,PrecioManual:netPrice(item),CodigoMonedaOrigen:item.currency,CodigoTipoIGVOrigen:'SIN_IGV',ZonasTexto:item.zones,MotivoPrecioManual:'Precio neto del detalle del presupuesto'}))field(row,name).value=value??'';
        if(!field(row,'TipoServicio').value)field(row,'TipoServicio').value=item.serviceType??'';field(row,'TipoServicio').title=field(row,'TipoServicio').value;
        if(!field(row,'NumeroNecesidad').value)field(row,'NumeroNecesidad').value=item.needNumber??'';
        field(row,'CodigoMaterialSAP').title=field(row,'CodigoMaterialSAP').value;row.querySelector('[data-line-material-code]').textContent=item.internalCode;
        sync(row);prices.set(row,new Map([[item.currency,field(row,'PrecioManual').value]]));
    }
    async function tariffFor(materialId,currency,tariffLineId){
        const params=new URLSearchParams({clientId:form.querySelector('[name="Draft.ClienteId"]').value||'0',date:header('FechaEmision').value,currency,materialId});
        if(tariffLineId)params.set('tariffLineId',tariffLineId);
        const response=await fetch('/Presupuestos/Tarifario?'+params,{headers:{Accept:'application/json'}});
        if(!response.ok)throw new Error();
        return (await response.json()).rows[0];
    }
    async function material(row,item){
        const selected=String(item.id);
        const currency=header('CodigoMoneda').value;
        for(const [name,value] of Object.entries({MaterialId:selected,PartidaTarifarioId:'',PrecioManual:'',Unidad:'',CodigoInterno:item.code,Descripcion:item.description,CodigoMaterialSAP:'',Posicion:'',ZonasTexto:''}))field(row,name).value=value;
        row.querySelector('[data-line-material-code]').textContent=item.code;
        const selection=new Map();
        prices.set(row,selection);
        try{
            const response=await fetch('/Maestros/materiales/'+encodeURIComponent(selected)+'/Predeterminados',{headers:{Accept:'application/json'}});
            if(!response.ok)throw new Error();
            const data=await response.json();
            if(prices.get(row)!==selection||!row.isConnected)return;
            field(row,'Unidad').value=data.fields.UnidadCodigo??'';
            field(row,'CodigoMaterialSAP').value=data.fields.CodigoMaterialSAP??'';
            field(row,'CodigoMaterialSAP').title=field(row,'CodigoMaterialSAP').value;
        }catch{
            if(prices.get(row)===selection)feedback.textContent='No se pudieron cargar los datos del material. Vuelve a seleccionarlo.';
        }
        if(prices.get(row)===selection&&header('CodigoMoneda').value===currency&&row.isConnected)await refreshPrice(row,currency,true);
        schedule();
    }
    async function refreshPrice(row,currency,selecting=false){
        const tariff=field(row,'PartidaTarifarioId').value,materialId=field(row,'MaterialId').value;
        const selection=prices.get(row);
        if(!selecting)field(row,'PrecioManual').value=selection.get(currency)??'';
        if(!materialId||field(row,'PrecioManual').value!=='')return;
        const current=()=>row.isConnected&&prices.get(row)===selection&&header('CodigoMoneda').value===currency&&field(row,'PrecioManual').value==='';
        try{
            const item=await tariffFor(materialId,currency,tariff);
            if(!current())return;
            if(item&&selecting)applyTariff(row,item);
            else if(item){
                field(row,'PrecioManual').value=netPrice(item);
                field(row,'PartidaTarifarioId').value=item.tariffLineId;
                selection.set(currency,field(row,'PrecioManual').value);
            }
            if(!item||item.price===null)feedback.textContent='El material no tiene precio en la moneda elegida. Registra el importe.';
        }catch{
            if(current())feedback.textContent='No se pudo cargar el precio. Vuelve a seleccionar el material o registra el importe.';
        }
        schedule();
    }
    function newRow(){
        if(body.children.length>=500){feedback.textContent='El presupuesto admite hasta 500 partidas.';return null;}
        const row=document.querySelector('#budget-line-template').content.firstElementChild.cloneNode(true);
        body.append(row);renumber();field(row,'CodigoMonedaOrigen').value=header('CodigoMoneda').value;field(row,'CodigoTipoIGVOrigen').value='SIN_IGV';
        field(row,'CentroSAPFICO').value=form.dataset.storeCenter??'';sync(row);return row;
    }
    function schedule(){if(!editable)return;clearTimeout(timer);timer=setTimeout(calculate,300);}
    async function calculate(){
        const current=++calculationRequest;renumber();
        const incomplete=[...body.querySelectorAll('[data-budget-line]')].some(row=>{const value=field(row,'Cantidad').value;return !/^[0-9]{1,13}(\.[0-9]{1,2})?$/.test(value)||Number(value)<=0;});
        if(incomplete){calculation.textContent='';body.querySelectorAll('[data-line-total]').forEach(element=>element.textContent='—');form.querySelectorAll('[data-budget-net],[data-budget-tax],[data-budget-total]').forEach(element=>element.textContent='—');return;}
        calculation.textContent='Calculando previsualización…';
        try{
            const response=await fetch('/Presupuestos/Calcular',{method:'POST',body:new FormData(form),headers:{Accept:'application/json'}});
            if(current!==calculationRequest)return;
            if(!response.ok||!response.headers.get('content-type')?.includes('application/json')){
                calculation.textContent=response.status===400?(await response.text()).replace(/^"|"$/g,''):'No se pudo calcular. Revisa los datos y vuelve a intentarlo.';
                form.querySelectorAll('[data-budget-net],[data-budget-tax],[data-budget-total]').forEach(element=>element.textContent='—');
                return;
            }
            const data=await response.json();if(current!==calculationRequest)return;
            body.querySelectorAll('[data-budget-line]').forEach((row,index)=>{
                row.querySelector('[data-line-total]').textContent=number(data.lines[index]?.net);
            });
            form.querySelector('[data-budget-net]').textContent=number(data.totals.net);
            form.querySelector('[data-budget-tax]').textContent=number(data.totals.tax);
            form.querySelector('[data-budget-total]').textContent=number(data.totals.total);
            calculation.textContent=data.totals.complete?'Importes calculados a dos decimales.':'Hay precios pendientes. Puedes guardar el borrador, pero no enviarlo hasta completarlos.';
        }catch{if(current===calculationRequest)calculation.textContent='No se pudo calcular. Puedes reintentar modificando el dato.';}
    }
    async function metadata(overwrite=false){
        const current=++metadataRequest;const parameters=new URLSearchParams({date:header('FechaEmision').value});
        const client=form.querySelector('[name="Draft.ClienteId"]').value,store=form.querySelector('[name="Draft.TiendaId"]').value;
        if(client)parameters.set('clientId',client);if(store)parameters.set('storeId',store);
        try{
            const response=await fetch('/Presupuestos/DatosIniciales?'+parameters,{headers:{Accept:'application/json'}});
            if(!response.ok||!response.headers.get('content-type')?.includes('application/json'))throw new Error();
            const data=await response.json();if(current!==metadataRequest)return;
            form.querySelector('[data-client-address]').textContent=data.Cliente?.DireccionFiscal||'Sin dirección registrada';
            form.querySelector('[data-store-address]').textContent=data.Tienda?.Direccion||'Sin dirección registrada';
            form.dataset.taxRate=data.Tasa?String(data.Tasa.Valor):'';
            form.dataset.storeCenter=data.Tienda?.CentroSAPFICO??'';
            const sap=header('CodigoSAPProveedor');
            if((sapPending||overwrite)&&!sap.value)sap.value=data.Proveedor?.CodigoSAPProveedor??'';
            sapPending=false;
            if(overwrite){
                for(const [key,value] of Object.entries({ZonaTexto:data.Tienda?.ZonaTexto,ContratoMarco:data.Proveedor?.ContratoMarco}))header(key).value=value??'';
                if(!header('Firmante').value)header('Firmante').value=data.Empresa?.NombreFirmante??'';
                if(!header('CargoFirmante').value)header('CargoFirmante').value=data.Empresa?.CargoFirmante??'';
            }
            feedback.textContent=data.Tasa?`IGV vigente en la fecha seleccionada: ${number(data.Tasa.Valor*100)} %. ${overwrite?'Datos sugeridos aplicados. Revisa los campos de impresión.':''}`:'No hay tasa de IGV configurada para esta fecha.';
            schedule();
        }catch{if(current===metadataRequest)feedback.textContent='No se pudieron consultar los datos sugeridos. Vuelve a intentarlo.';}
    }
    function close(){if(!busy){lookupRequest++;setDialog(dialog,false);}}
    async function load(){
        const current=++lookupRequest;busy=true;setDialog(dialog,true,true,close);lookupBody.replaceChildren();lookupStatus.textContent='Cargando…';previous.disabled=true;next.disabled=true;
        try{
            const parameters=new URLSearchParams({search:query.elements.search.value,page:String(page)});let endpoint;
            endpoint='/Maestros/'+(lookupMode==='material'?'materiales':lookupMode==='service'?'tipos-servicio':reference.dataset.entity)+'/Ayuda';
            if(reference?.dataset.entity==='tiendas')parameters.set('parentId',form.querySelector('[name="Draft.ClienteId"]').value);
            const response=await fetch(endpoint+'?'+parameters,{headers:{Accept:'application/json'}});
            if(!response.ok||!response.headers.get('content-type')?.includes('application/json'))throw new Error();
            const data=await response.json();if(current!==lookupRequest)return;
            for(const item of data.rows){
                const tr=document.createElement('tr');
                const values=[item.code,item.description,item.extra??''];
                for(const value of values){const td=document.createElement('td');td.textContent=value;tr.append(td);}
                const cell=document.createElement('td'),button=document.createElement('button');button.type='button';button.className='hyt-button-secondary';button.textContent='Seleccionar';
                button.addEventListener('click',()=>{
                    if(lookupMode==='material'){material(rowTarget,item);}else if(lookupMode==='service'){field(rowTarget,'TipoServicio').value=item.description;field(rowTarget,'TipoServicio').title=item.description;}else{
                        const previousId=reference.querySelector('input').value;
                        reference.querySelector('input').value=['zonas','especialidades'].includes(reference.dataset.entity)?item.description:String(item.id);reference.querySelector('[data-reference-label]').textContent=item.code+' · '+item.description;
                        if(reference.dataset.entity==='clientes'&&previousId!==String(item.id)){
                            form.querySelector('[data-client-address]').textContent='Consultando dirección…';form.querySelector('[data-store-address]').textContent='Selecciona una tienda';
                            const store=form.querySelector('[data-budget-reference="TiendaId"]');store.querySelector('input').value='';store.querySelector('[data-reference-label]').textContent='Sin seleccionar';
                            sapPending=true;
                            for(const key of ['CodigoSAPProveedor','ContratoMarco','ZonaTexto'])header(key).value='';form.dataset.storeCenter='';
                        }
                        if(reference.dataset.entity==='tiendas')form.querySelector('[data-store-address]').textContent='Consultando dirección…';
                        if(['clientes','tiendas'].includes(reference.dataset.entity))metadata();
                    }
                    close();schedule();
                });cell.append(button);tr.append(cell);lookupBody.append(tr);
            }
            lookupStatus.textContent=data.rows.length?`${data.total} registros encontrados.`:'No se encontraron registros con los filtros seleccionados.';
            dialog.querySelector('[data-budget-page]').textContent=`Página ${page} · 20 por página`;previous.disabled=page<=1;next.disabled=page*20>=data.total;
        }catch{if(current===lookupRequest)lookupStatus.textContent='No se pudo cargar la ayuda. Revisa los datos seleccionados y vuelve a buscar.';}
        finally{if(current===lookupRequest){busy=false;setDialog(dialog,true,false,close);}}
    }
    function open(mode,target=null){
        if(!editable)return;
        
        if(mode==='reference'&&target.dataset.entity==='tiendas'&&!form.querySelector('[name="Draft.ClienteId"]').value){feedback.textContent='Selecciona primero el cliente.';return;}
        lookupMode=mode;reference=mode==='reference'?target:null;rowTarget=['material','service'].includes(mode)?target:null;page=1;query.reset();
        dialog.querySelector('#budget-lookup-title').textContent=mode==='material'?'Seleccionar material':mode==='service'?'Seleccionar tipo de servicio':'Seleccionar '+({tiendas:'tienda',clientes:'cliente',zonas:'zona',especialidades:'especialidad'}[target.dataset.entity]||'registro');
        const labels=['Código','Descripción','Información','Acción'];
        const head=dialog.querySelector('[data-budget-lookup-head]'),tr=document.createElement('tr');head.replaceChildren();for(const label of labels){const th=document.createElement('th');th.textContent=label;tr.append(th);}head.append(tr);load();
    }
    renumber();body.querySelectorAll('[data-budget-line]').forEach(sync);
    if(editable){
        form.querySelector('[data-manual-line]').addEventListener('click',()=>{const row=newRow();if(row){field(row,'Descripcion').focus();schedule();}});
        form.querySelectorAll('[data-reference-open]').forEach(button=>button.addEventListener('click',()=>open('reference',button.closest('[data-budget-reference]'))));
        form.querySelectorAll('[data-budget-clear-reference]').forEach(button=>button.addEventListener('click',()=>{const box=button.closest('[data-budget-reference]');box.querySelector('input').value='';box.querySelector('[data-reference-label]').textContent='Sin seleccionar';}));
        body.addEventListener('click',event=>{
            const remove=event.target.closest('[data-remove-line]');if(remove){remove.closest('[data-budget-line]').remove();renumber();schedule();}
            const chooseMaterial=event.target.closest('[data-line-material]');if(chooseMaterial)open('material',chooseMaterial.closest('[data-budget-line]'));
            const service=event.target.closest('[data-line-service]');if(service)open('service',service.closest('[data-budget-line]'));
        });
        body.addEventListener('input',event=>{const row=event.target.closest('[data-budget-line]');if(row){sync(row);if(event.target.matches('[data-origin-price]'))prices.get(row).set(field(row,'CodigoMonedaOrigen').value,field(row,'PrecioManual').value);}});
        header('CodigoMoneda').addEventListener('change',()=>{
            const currency=header('CodigoMoneda').value;
            for(const row of body.querySelectorAll('[data-budget-line]')){
                const cache=prices.get(row);cache.set(field(row,'CodigoMonedaOrigen').value,field(row,'PrecioManual').value);
                field(row,'CodigoMonedaOrigen').value=currency;refreshPrice(row,currency);sync(row);
            }
            feedback.textContent='Revisa precios en la nueva moneda. Los precios se consultan en la moneda seleccionada.';schedule();
        });
        header('FechaEmision').addEventListener('change',()=>metadata(false));
        form.addEventListener('input',schedule);form.addEventListener('change',schedule);form.addEventListener('submit',renumber);
        if(body.children.length)schedule();
    }
    dialog.querySelectorAll('[data-budget-close]').forEach(button=>button.addEventListener('click',close));
    query.addEventListener('submit',event=>{event.preventDefault();if(!busy){page=1;load();}});
    previous.addEventListener('click',()=>{if(!busy&&page>1){page--;load();}});next.addEventListener('click',()=>{if(!busy){page++;load();}});
}
