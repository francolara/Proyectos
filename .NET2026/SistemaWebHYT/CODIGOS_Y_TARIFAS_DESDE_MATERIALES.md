# Códigos internos y tarifas desde materiales

Actualización: 05/11/2026. Autor: FRANCO LARA.

## Funcionamiento

- Clientes, tiendas, zonas, especialidades, tipos de servicio, materiales y tarifarios reciben un código interno al guardar. El servidor ignora cualquier código enviado desde el navegador y conserva el original al editar. Los registros históricos no se renumeran.
- Los otros mantenimientos muestran su identidad interna de solo lectura. Los números internos de los documentos conservan su generación automática. Las referencias externas del cliente y el código SAP siguen siendo editables.
- Unidades de medida muestra el identificador interno separado de **Abreviatura**, que admite UND, KG u otra abreviatura única. La columna física Codigo conserva la abreviatura para mantener compatibilidad con los presupuestos y reportes existentes.
- El maestro de materiales ya no solicita Posición. Este dato sigue disponible en cada partida del tarifario. La columna antigua de materiales se conserva sin uso para no eliminar información histórica.
- Especialidades y Tipos de servicio tienen entradas propias en el menú Maestros.

## Tarifarios en la ficha del material

Guardar primero un material nuevo. En su ficha aparece **Tarifarios del material**, con veinte filas por página desde SQL. Se requiere consulta de materiales y tarifarios; agregar o editar requiere edición de tarifarios.

**Agregar tarifario** abre una ventana con solo cuatro campos:

- Tipo de IGV: CON IGV o SIN IGV.
- Precio en soles.
- Precio en dólares, opcional y registrado manualmente.
- Estado: activo o inactivo.

Se debe ingresar al menos un precio en soles o dólares; pueden registrarse ambos. Los precios admiten hasta dos decimales. No hay selección de versiones, pasos de publicación ni conversión automática de monedas. El material queda asociado automáticamente. Guardar incorpora el registro para su uso en presupuestos si está activo.

Cada registro conserva su usuario y fecha de creación y actualización. **Ver** y **Editar** abren la misma ventana; las fechas y responsables propios aparecen en consulta y edición. Las tarifas convencionales previamente publicadas se conservan protegidas.

El listado muestra IGV, precios, estado, responsables y fechas. Guardar un tarifario actualiza el listado y conserva los campos pendientes del material. No guarda automáticamente esos cambios pendientes.

## Base local y entrega

Aplicados Script/023_TarifarioMaterialDirecto.sql, spu_MaterialTarifa_Guardar y los procedimientos de consulta y permisos relacionados en SistemaWebHYT local. Los 84 procedimientos instalados coinciden con sus fuentes. Se actualizan los paquetes SQL y sus huellas. No se insertaron datos ficticios en la base de trabajo.

## Validación

- Solución y publicación Release: sin errores ni advertencias.
- 173 comprobaciones SQL aprobadas en una base temporal: alta directa, disponibilidad en presupuestos, estado, auditoría, concurrencia y protección de precios convencionales publicados.
- Pruebas HTTP aprobadas: validaciones, permisos, pertenencia del material y compatibilidad de los accesos anteriores.
- Chrome: cuatro controles, alta con USD vacío, edición de precios y estado, conservación de cambios pendientes del material.
- 255 comprobaciones de interfaz en claro/oscuro, escritorio/tablet/móvil, popups y contraste.
- Validación SQL: 155 fuentes, 48 tablas y 108 relaciones.

Para probar en Visual Studio, iniciar nuevamente GestionServicios.Web con F5 después de esta actualización.

## Listado general de Tarifarios

El menú Tarifarios muestra también los precios agregados desde materiales. Permite buscar por código, descripción o SAP del material, con veinte filas por página. Nuevo registro permite elegir el material desde la ayuda y completar IGV, soles y/o dólares y estado. Ambos accesos comparten el mismo registro y sus fechas/responsables; no es necesario volver a registrar los precios ya existentes. Validado con 177 comprobaciones SQL, pruebas HTTP y Chrome de alta y listado unificado en escritorio/móvil.
