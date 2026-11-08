// Firma: FRANCO LARA - 06/11/2026 | Nombres comerciales de estados del presupuesto/OC y navegación por fuente independiente. Curvas con nombres legibles Confirmación del cliente y Facturado sobre registros vigentes.
// Firma: FRANCO LARA - 03/11/2026 | Gráficos SVG HYT con curvas sin sobrepasar valores, detalle por teclado y agrupación acotada.
// Firma: FRANCO LARA - 05/11/2026 | Agrupación día/semana/mes, eventos diferenciados y navegación por intervalo con espera compartida y gráficos legibles en ambos temas.
(()=>{
const source=document.querySelector('#hyt-dashboard-data');if(!source)return;
const model=JSON.parse(source.textContent),data=model.data,ns='http://www.w3.org/2000/svg';
const stateLabel=state=>({BORRADOR:'Registrado',ENVIADO:'Enviado',OBSERVADO:'Observado',APROBADO:'Aceptada',RECHAZADO:'Rechazada',CANCELADO:'Cancelado',REGISTRADA:'Registrada',TRABAJO_COMPLETADO:'Completado',CERRADA_OPERATIVAMENTE:'Cerrada operativamente',CANCELADA:'Cancelada'}[state]??state);
const colors=['#c91620','var(--hyt-chart-ink)' ,'var(--hyt-chart-muted)' ];
const value=n=>Number.isFinite(Number(n))?Math.max(0,Number(n)):0;
function node(tag,attrs={},text){const element=document.createElementNS(ns,tag);for(const [key,val] of Object.entries(attrs))element.setAttribute(key,String(val));if(text!==undefined)element.textContent=text;return element;}
function title(element,text){element.append(node('title',{},text));return element;}
const day=iso=>Date.parse(iso+'T00:00:00Z'),date=ms=>new Intl.DateTimeFormat('es-PE',{day:'2-digit',month:'2-digit',year:'numeric',timeZone:'UTC'}).format(ms);
const from=day(model.from),until=day(model.to),days=Math.round((until-from)/86400000)+1;
const series=['Presupuestos','Ordenes','Confirmaciones','Marcados'],seriesLabels={Presupuestos:'Presupuestos',Ordenes:'OC del cliente',Confirmaciones:'Confirmación del cliente',Marcados:'Facturado'};
let buckets=[],step=1,grouping='day';
function regroup(){
 grouping=document.querySelector('[data-chart-group]')?.value||'day';buckets=[];
 if(grouping==='month'){
  const start=new Date(from),end=new Date(until),firstMonth=start.getUTCFullYear()*12+start.getUTCMonth(),lastMonth=end.getUTCFullYear()*12+end.getUTCMonth(),stride=Math.max(1,Math.ceil((lastMonth-firstMonth+1)/90));
  for(let month=firstMonth;month<=lastMonth;month+=stride){const startDate=new Date(0),endDate=new Date(0);startDate.setUTCFullYear(Math.floor(month/12),month%12,1);endDate.setUTCFullYear(Math.floor((month+stride)/12),(month+stride)%12,1);buckets.push({date:Math.max(from,startDate.getTime()),end:Math.min(until,endDate.getTime()-86400000),Presupuestos:0,Ordenes:0,Confirmaciones:0,Marcados:0});}step=0;
 }else{
  const unit=grouping==='week'?7:1;step=Math.max(unit,Math.ceil(days/90/unit)*unit);
  buckets=Array.from({length:Math.ceil(days/step)},(_,index)=>({date:from+index*step*86400000,end:Math.min(until,from+((index+1)*step-1)*86400000),Presupuestos:0,Ordenes:0,Confirmaciones:0,Marcados:0}));
 }
 for(const record of data.Movimientos){const time=day(record.Fecha),bucket=buckets.find(item=>time>=item.date&&time<=item.end);if(bucket)for(const key of series)bucket[key]+=value(record[key]);}
}
regroup();
const trend=document.querySelector('[data-chart=trend]'),readout=document.querySelector('[data-chart-readout]'),note=document.querySelector('[data-chart-note]');
const controls=document.querySelector('.hyt-chart-controls');
let selector;
if(controls){const label=document.createElement('label');label.textContent='Detalle por fecha ';selector=document.createElement('select');selector.setAttribute('aria-label','Consultar intervalo del gráfico');label.append(selector);controls.append(label);}
const labelFor=bucket=>date(bucket.date)+(bucket.end>bucket.date?' – '+date(bucket.end):'');
function read(index){const bucket=buckets[index];if(!bucket)return;const selected=[...document.querySelectorAll('[data-chart-series]:checked')].map(input=>input.dataset.chartSeries);readout.textContent=labelFor(bucket)+' · '+selected.map(key=>seriesLabels[key]+': '+bucket[key].toLocaleString('es-PE')).join(' · ');}
function renderTrend(){
if(!trend)return;const linked=document.querySelector('[data-chart-link]');if(linked)linked.hidden=true;trend.replaceChildren();selector?.replaceChildren();
const keys=[...document.querySelectorAll('[data-chart-series]:checked')].map(input=>input.dataset.chartSeries);
const hasRecords=buckets.some(bucket=>keys.some(key=>bucket[key]>0));
trend.style.display=keys.length&&hasRecords?'':'none';if(selector)selector.disabled=!keys.length||!hasRecords;
note.textContent=!keys.length?'Selecciona una serie para consultar los registros.':!hasRecords?'No hay registros de las series seleccionadas en este período.':grouping==='month'?'Intervalos por mes, hasta 90 grupos; detalle exacto al seleccionar.':step>1?`Cada punto reúne ${step} días desde el inicio del período; hasta 90 intervalos.`:'Cada punto representa un día, incluidos los días sin registros.';
if(!keys.length||!hasRecords){readout.textContent='';return;}
const maximum=Math.max(4,Math.ceil(Math.max(0,...buckets.flatMap(bucket=>keys.map(key=>bucket[key])))/4)*4),left=52,right=740,top=18,bottom=220;
const x=index=>buckets.length===1?(left+right)/2:left+index*(right-left)/(buckets.length-1),y=count=>bottom-count/maximum*(bottom-top);
for(let tick=0;tick<=4;tick++){const count=maximum*tick/4,position=y(count);trend.append(node('line',{x1:left,x2:right,y1:position,y2:position,stroke:'var(--hyt-border)'}));trend.append(node('text',{x:42,y:position+4,'text-anchor':'end'},Math.ceil(count).toLocaleString('es-PE')));}
for(const key of keys){
const color=colors[key==='Presupuestos'||key==='Confirmaciones'?0:1],points=buckets.map((bucket,index)=>[x(index),y(bucket[key])]);let path=`M ${points[0][0]} ${points[0][1]}`;
for(let index=1;index<points.length;index++){const previous=points[index-1],next=points[index],middle=(previous[0]+next[0])/2;path+=` C ${middle} ${previous[1]},${middle} ${next[1]},${next[0]} ${next[1]}`;}
trend.append(node('path',{d:path,fill:'none',stroke:color,'stroke-width':3,'stroke-linecap':'round','stroke-dasharray':key==='Confirmaciones'?'8 5':key==='Marcados'?'3 5':'none'}));
buckets.forEach((bucket,index)=>{const text=labelFor(bucket)+' · '+seriesLabels[key]+': '+bucket[key],circle=title(node('circle',{cx:x(index),cy:y(bucket[key]),r:4,fill:color,tabindex:0,role:'img','aria-label':text}),text);circle.addEventListener('focus',()=>{if(selector)selector.value=String(index);read(index);});circle.addEventListener('pointerenter',()=>read(index));const select=()=>{read(index);const link=document.querySelector('[data-chart-link]');if(!link)return;link.hidden=!['Presupuestos','Ordenes'].includes(key);if(link.hidden)return;const url=new URL(key==='Presupuestos'?'/Presupuestos':'/Seguimiento/oc-cliente',location.origin);url.searchParams.set('from',new Date(bucket.date).toISOString().slice(0,10));url.searchParams.set('to',new Date(bucket.end).toISOString().slice(0,10));if(model.client)url.searchParams.set('client',model.client);if(model.store)url.searchParams.set('store',model.store);link.href=url.pathname+url.search;link.textContent='Abrir '+seriesLabels[key]+' del intervalo';};circle.addEventListener('click',select);circle.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();select();}});trend.append(circle);});
}
trend.append(node('text',{x:left,y:252},date(from)),node('text',{x:right,y:252,'text-anchor':'end'},date(until)));
buckets.forEach((bucket,index)=>{const option=document.createElement('option');option.value=String(index);option.textContent=labelFor(bucket);selector?.append(option);});if(selector){selector.value=String(buckets.length-1);}read(buckets.length-1);
}
document.querySelector('[data-chart-group]')?.addEventListener('change',()=>{regroup();renderTrend();const link=document.querySelector('[data-chart-link]');if(link)link.hidden=true;});
selector?.addEventListener('change',()=>read(Number(selector.value)));document.querySelectorAll('[data-chart-series]').forEach(input=>input.addEventListener('change',renderTrend));renderTrend();
for(const states of document.querySelectorAll('[data-state-source]')){
const isBudget=states.dataset.stateSource==='PRESUPUESTO',rows=isBudget?data.EstadosPresupuesto:data.EstadosOC;
if(states){states.replaceChildren();if(!rows.length){states.style.display='none';const empty=document.createElement('p');empty.textContent='Sin documentos para distribuir por estado.';states.after(empty);}else{
const maximum=Math.max(1,...rows.map(row=>value(row.Cantidad))),left=40,width=700,bottom=220;
for(let tick=0;tick<=4;tick++){const position=bottom-tick*48;states.append(node('line',{x1:left,x2:750,y1:position,y2:position,stroke:'var(--hyt-border)'}));}
rows.forEach((row,index)=>{const slot=width/rows.length,x=left+slot*index+slot*.2,height=value(row.Cantidad)/maximum*180,label=stateLabel(row.Estado),text=label+': '+row.Cantidad;const bar=title(node('rect',{x,y:bottom-height,width:slot*.6,height,rx:3,fill:colors[0],tabindex:0,role:'link','aria-label':text+' · Abrir listado'}),text);const open=()=>{const url=new URL(isBudget?'/Presupuestos':'/Seguimiento/oc-cliente',location.origin);url.searchParams.set('from',model.from);url.searchParams.set('to',model.to);url.searchParams.set('state',row.Estado);if(model.client)url.searchParams.set('client',model.client);if(model.store)url.searchParams.set('store',model.store);globalThis.hytLoading?.navigation();location.assign(url.pathname+url.search);};bar.addEventListener('click',open);bar.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();open();}});states.append(bar);states.append(node('text',{x:x+slot*.3,y:bottom-height-10,'text-anchor':'middle'},row.Cantidad));states.append(title(node('text',{x:x+slot*.3,y:244,'text-anchor':'middle',transform:`rotate(-18 ${x+slot*.3} 244)`},label),label));});
}}
if(states&&rows.length){const legend=document.createElement('ul');legend.className='hyt-states-legend';for(const row of rows){const item=document.createElement('li');item.textContent=stateLabel(row.Estado)+' · '+value(row.Cantidad).toLocaleString('es-PE');legend.append(item);}states.after(legend);}
}
const donut=document.querySelector('[data-chart=donut]'),legend=document.querySelector('[data-donut-legend]');
if(donut){const categories=[{label:'Sin confirmar',count:value(data.SinConfirmar),color:colors[2]},{label:'Confirmadas sin marcar',count:value(data.ConfirmadasSinMarcar),color:colors[1]},{label:'Facturadas',count:value(data.Facturadas),color:colors[0]}],total=categories.reduce((sum,item)=>sum+item.count,0),circumference=2*Math.PI*84;let offset=0;
donut.append(node('circle',{cx:180,cy:125,r:84,fill:'none',stroke:'var(--hyt-border)','stroke-width':28}));
for(const item of categories){if(total&&item.count){
const fraction=item.count/total,text=item.label+': '+item.count+' ('+(fraction*100).toFixed(2)+' %)',a=offset/circumference*Math.PI*2-Math.PI/2,b=a+fraction*Math.PI*2;
const point=(radius,angle)=>`${180+radius*Math.cos(angle)} ${125+radius*Math.sin(angle)}`;
let segment;
if(item.count===total)segment=node('circle',{cx:180,cy:125,r:84,fill:'none',stroke:item.color,'stroke-width':28});
else segment=node('path',{d:`M ${point(98,a)} A 98 98 0 ${fraction>.5?1:0} 1 ${point(98,b)} L ${point(70,b)} A 70 70 0 ${fraction>.5?1:0} 0 ${point(70,a)} Z`,fill:item.color});
segment.setAttribute('tabindex','0');segment.setAttribute('role','img');segment.setAttribute('aria-label',text);segment.setAttribute('data-segment','');donut.append(title(segment,text));offset+=fraction*circumference;}
const line=document.createElement('li'),swatch=document.createElement('span');swatch.className='hyt-chart-swatch';swatch.style.backgroundColor=item.color;line.append(swatch,document.createTextNode(item.label+' · '+item.count.toLocaleString('es-PE')));legend?.append(line);}
const center=node('text',{x:180,y:130,'text-anchor':'middle'},total.toLocaleString('es-PE'));center.style.fontSize='32px';center.style.fill='var(--hyt-chart-ink)' ;donut.append(center,node('text',{x:180,y:154,'text-anchor':'middle'},total?'OC activas':'Sin OC activas'));
}
})();