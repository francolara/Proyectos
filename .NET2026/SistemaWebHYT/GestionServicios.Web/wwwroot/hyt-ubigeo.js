// Firma: FRANCO LARA - 05/11/2026 | Ubigeo dependiente departamento, provincia y distrito sin perder selección al editar.
for(const box of document.querySelectorAll('[data-ubigeo]')){
 const department=box.querySelector('[data-ubigeo-level=Departments]'),province=box.querySelector('[data-ubigeo-level=Provinces]'),district=box.querySelector('[data-ubigeo-level=Districts]');
 const provinces=[...province.options].map(o=>({value:o.value,text:o.textContent})),districts=[...district.options].map(o=>({value:o.value,text:o.textContent}));
 const fill=(select,options,prefix,value)=>{select.replaceChildren(new Option('Selecciona',''));for(const item of options)if(item.value&&prefix&&item.value.startsWith(prefix))select.add(new Option(item.text,item.value));select.value=value;};
 const address=box.closest('form').querySelector('[name="Fields[DireccionFiscal]"],[name="Fields[Direccion]"]');
 const required=()=>{district.required=!!address?.value.trim();};address?.addEventListener('input',required);required();
 const initial=box.dataset.value||'';department.value=initial.slice(0,2);fill(province,provinces,department.value,initial.slice(0,4));fill(district,districts,province.value,initial);
 department.addEventListener('change',()=>{fill(province,provinces,department.value,'');fill(district,districts,'','');});
 province.addEventListener('change',()=>fill(district,districts,province.value,''));
}
