# Análisis de FrmBusqueda

Fecha: 15/10/2026

## Responsabilidad

`Formularios/FrmBusqueda.frm` es la ayuda genérica del sistema. Recibe un tipo de ayuda, muestra los campos uniformes `cod` y `des`, y retorna el registro elegido a la ventana que la invocó.

## Funcionamiento anterior

`setSql` contiene 88 casos declarados (87 tipos únicos). Construye un `SELECT` según el tipo recibido; `fill` reemplaza `%X%` por el texto escrito, concatena `SqlAdic` y abre el recordset. Después de la primera búsqueda, el cambio de cada carácter vuelve a consultar la base.

`SqlAdic` puede contener filtros SQL construidos por los mantenimientos. Por compatibilidad, esos casos se conservan temporalmente en el camino anterior hasta reemplazarlos por parámetros explícitos.

La búsqueda ya no se ejecuta por cada cambio de `TxtBusq`. La carga inicial se conserva y las consultas posteriores se realizan solo al presionar Enter; así se evita una llamada por cada carácter contra la base en la nube.

La configuración de visualización de clientes por vendedor se carga una sola vez durante el login. `FrmBusqueda` utiliza las variables globales resultantes y ya no consulta `Parametros` ni `Vendedores` cada vez que abre una ayuda.

## Primera migración

Se creó `dbo.spu_Ayuda_PersonasComercial` y `FrmBusqueda` lo usa cuando no recibe `SqlAdic` para estas ayudas frecuentes:

- Clientes, proveedores, vendedores, usuarios, choferes y empresas de transporte.
- Productos, almacenes, centros de costo, listas de precios y vehículos.

El procedimiento recibe tipo, texto, empresa, sucursal, usuario y el indicador de restricción de clientes por vendedor. Devuelve siempre `cod` y `des`; no recibe SQL dinámico.

Se agregó `dbo.spu_Ayuda_Generales` para filtros simples de geografía, almacén de venta, motivos, tipo de movimiento de caja, niveles, presentaciones y formas de pago por cliente. `FrmBusqueda` extrae únicamente valores de condiciones simples de igualdad, como país, departamento, provincia, sucursal, documento, indicador de ingreso/salida, tipo de nivel, nivel predecesor, producto o cliente, y los transmite como parámetros. Las condiciones complejas continúan temporalmente en la ruta anterior.

Se agregó `dbo.spu_Ayuda_Catalogos` para 13 ayudas que leen la tabla `Datos`: tipo de persona, país, tipo de producto, marca de vehículo, tipo de ticket, configuraciones de letra/cobro/efectivo, tipos de nivel y conceptos de costeo. También atiende moneda activa, unidades de medida, tipos y formas de pago, documentos, unidades de producción y direcciones de recojo. Solo se usa cuando no existe `SqlAdic`.

`dbo.spu_Ayuda_Generales` también cubre almacenes y sucursales de destino. Para no alterar el resultado heredado, estas dos ayudas no filtran por el texto de búsqueda, pues su `SELECT` anterior tampoco lo hacía.

El bloque administrativo simple también se trasladó a `dbo.spu_Ayuda_Catalogos`: marcas, conceptos de ingreso/salida, perfiles, empresas, tipos de nivel, cajas, permisos, cuentas corrientes, bancos, grupos de producto, zonas, ubicaciones, tallas/pesos, canales y tipos de documento de identidad.

Las ayudas de personas sin `SqlAdic` también se resuelven en `dbo.spu_Ayuda_PersonasComercial`: camales, cobradores, avales, vendedores jefe, sucursales y personas disponibles para ser cliente, proveedor, vendedor, usuario, chofer, empresa de transporte, cobrador o aval. Las exclusiones usan `NOT EXISTS` con empresa tipada.

Se incorporaron además las cajas asignadas al usuario y los estados de documento a `dbo.spu_Ayuda_Catalogos`.

`dbo.spu_Ayuda_Generales` incorpora documentos de exportación y contactos de clientes/proveedores. El documento origen y el cliente se extraen de condiciones simples y se transmiten como parámetros.

## Próximas etapas

1. Convertir los filtros de `SqlAdic` en parámetros tipados por cada ayuda.
2. Migrar los demás catálogos de `setSql` por grupos funcionales.
3. Evaluar si la carga inicial de cada ayuda debe limitarse cuando los catálogos crezcan, sin alterar las ayudas que requieren explorar toda la lista.

## Exclusiones confirmadas

No se migrarán `CENTROCOSTOCLIENTE` ni `CAJASUSUARIOFILTRO`, porque ya no se usan. Sus llamadas actuales quedan fuera del alcance de esta optimización.
