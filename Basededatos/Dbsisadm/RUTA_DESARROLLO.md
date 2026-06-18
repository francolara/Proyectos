# Ruta de desarrollo recomendada

## Conceptos rescatados del sistema legado

1. El sistema debe ser multiempresa, pero controlado por una cuenta administradora asociada al usuario autenticado.
2. La contabilidad no debe vivir aislada: depende de maestros compartidos como empresa, clientes, proveedores, moneda y tipo de cambio.
3. El asiento contable es el núcleo transversal.
4. Compras, ventas y mantenimientos deben terminar generando o alimentando asientos.
5. El origen contable es una pieza clave para separar asientos manuales, compras, ventas, bancos, cierres y otros procesos.

## Modelo objetivo

- Aplicacion web: ASP.NET Core MVC
- Seguridad: ASP.NET Identity
- Acceso a datos: ADO.NET + Stored Procedures
- Base de datos unica: SQL Server `Dbsisadm`
- Multiempresa: un usuario administra una cuenta suscriptora y dentro de ella puede operar varias empresas; la sesion trabaja con una empresa activa

## Base ya definida en esta etapa

### Seguridad y contexto

- `SEG_Empresa`
- `SEG_CuentaAdministradora`
- `SEG_UsuarioCuentaAdministradora`
- `SEG_CuentaAdministradoraSuscripcion`
- `SEG_CuentaAdministradoraSuscripcionMovimiento`
- `SEG_CuentaAdministradoraSuscripcionPago`
- `SEG_UsuarioEmpresa`
- `usp_SEG_RegistrarCuentaAdministradoraConEmpresa`
- `usp_SEG_RegistrarEmpresaCuentaAdministradora`
- `usp_SEG_ListarCuentasAdministradorasSuscripcion`
- `usp_SEG_ActualizarSuscripcionCuentaAdministradora`
- `usp_SEG_ObtenerContextoSuscripcionPorEmpresa`
- `usp_SEG_ListarEmpresasPorUsuario`
- `usp_SEG_AsignarUsuarioEmpresa`

### Maestros administrativos

- `ADM_Persona`
- `ADM_Cliente`
- `ADM_Proveedor`
- `ADM_Moneda`
- `ADM_TipoCambio`
- `ADM_TipoComprobante`
- `UbigeoDepartamentos`
- `UbigeoProvincias`
- `UbigeoDistritos`
- `TiposDocumentoIdentidadSunat`
- `usp_ADM_ListarProveedoresActivosPorEmpresa`
- `usp_ADM_ListarClientesActivosPorEmpresa`
- `usp_ADM_ListarTiposComprobanteActivos`
- `usp_ADM_ListarTiposDocumentoIdentidadSunat`
- `usp_ADM_ListarUbigeoDepartamentos`
- `usp_ADM_ListarUbigeoProvincias`
- `usp_ADM_ListarUbigeoDistritos`
- `usp_ADM_ListarPersonasPorEmpresa`
- `usp_ADM_ObtenerPersona`
- `usp_ADM_GuardarPersona`

### Nucleo contable inicial

- `CON_PlanCuenta`
- `CON_Origen`
- `CON_Asiento`
- `CON_AsientoDetalle`
- `CON_CorrelativoAsiento`
- `CON_CuentaDestinoRegla`
- `CON_CuentaDestinoReglaDetalle`
- `CON_ConfiguracionContabilizacion`
- `CON_ConfiguracionContabilizacionDetalle`
- `COM_Compra`
- `COM_CompraDetalle`
- `VEN_Venta`
- `VEN_VentaDetalle`
- `usp_CON_ListarOrigenesActivos`
- `usp_CON_GuardarOrigenPorEmpresa`
- `usp_CON_ListarPlanCuentaPorEmpresa`
- `usp_CON_GuardarPlanCuentaPorEmpresa`
- `usp_ADM_ListarMonedasActivas`
- `usp_CON_ListarConfiguracionContabilizacionPorEmpresa`
- `usp_CON_ObtenerConfiguracionContabilizacion`
- `usp_CON_GuardarConfiguracionContabilizacion`
- `usp_CON_EliminarConfiguracionContabilizacion`
- `usp_CON_ListarAsientosPorEmpresa`
- `usp_CON_ObtenerAsiento`
- `usp_CON_GuardarAsientoManual`
- `usp_COM_ListarComprasPorEmpresa`
- `usp_COM_ObtenerCompra`
- `usp_COM_GuardarCompraConAsiento`
- `usp_VEN_ListarVentasPorEmpresa`
- `usp_VEN_ObtenerVenta`
- `usp_VEN_GuardarVentaConAsiento`
- `usp_CON_GenerarOrigenesBaseEmpresa`
- `usp_CON_ObtenerSiguienteNumeroAsiento`
- `usp_CON_ListarCuentasDestinoReglaPorEmpresa`
- `usp_CON_ObtenerCuentaDestinoRegla`
- `usp_CON_GuardarCuentaDestinoRegla`
- `usp_CON_EliminarCuentaDestinoRegla`

## Secuencia recomendada de desarrollo

### Fase 1. Seguridad y multiempresa

- Login con Identity
- Registro de cuenta administradora con empresa inicial
- Control comercial de suscripcion por cuenta administradora
- Seleccion de empresa activa por sesion
- Restriccion de consultas y grabaciones por `IdEmpresa`

### Fase 2. Maestros base

- Personas
- Tipos de documento SUNAT
- Ubigeo por departamento, provincia y distrito
- Clientes
- Proveedores
- Monedas
- Tipo de cambio

### Fase 3. Contabilidad estructural

- Plan de cuentas por empresa
- Origenes contables
- Reglas de cuentas destino por ejercicio para distribuciones tipo 6/79 u otras equivalencias
- Configuracion contable automatica por modulo y escenario para compras y ventas
- Asiento manual
- Validacion de cuadre Debe/Haber
- Numeracion de asientos por empresa, origen y periodo mensual con reinicio en cada mes

### Fase 4. Mantenimientos contables

- Ejercicio contable
- Periodos contables
- Configuracion de cuentas por defecto
- Parametros de compras y ventas para contabilizacion automatica segun escenario como mercaderia, gasto o servicio

### Fase 5. Registro de compras

- Provisiones de compras
- Relacion con proveedor
- Moneda y tipo de cambio
- Generacion automatica de asiento por origen `COM`
- Aplicacion de configuracion contable segun escenario como mercaderia, gasto o servicio

### Fase 6. Registro de ventas

- Comprobantes de venta
- Relacion con cliente
- Moneda y tipo de cambio
- Generacion automatica de asiento por origen `VEN`
- Aplicacion de configuracion contable segun escenario como mercaderia, servicio u otros modelos de venta

### Fase 7. Caja, bancos y cancelaciones

- Ingresos y egresos
- Aplicacion contra documentos
- Asientos de tesoreria

### Fase 8. Reportes y cierres

- Libro diario
- Mayor
- Balance de comprobacion
- Estado de resultados
- Balance general
- Cierre mensual y anual

## Criterios tecnicos

1. Toda tabla funcional debe incluir `IdEmpresa` cuando corresponda aislamiento multiempresa.
2. Todo modulo transaccional debe terminar en un Stored Procedure de grabacion.
3. La UI no debe contener reglas contables complejas.
4. La numeracion, validacion y reglas de cuadre deben quedar en SQL Server o en la capa de negocio ADO.NET, no en JavaScript.
5. El plan de cuentas debe permitir niveles y cuentas de movimiento.
6. El correlativo funcional del comprobante debe ser unico por `IdEmpresa + IdOrigen + Periodo + NumeroAsiento`.
7. Las reglas de cuenta destino deben versionarse por ejercicio y quedar separadas del plan de cuentas base.
8. Todo listado web de mantenimiento o registro debe paginarse desde Stored Procedure en bloques de 20 registros.
9. Los filtros de texto deben resolverse en SQL Server para evitar cargas completas antes de la consulta.
10. Los modulos por periodo como asientos, compras y ventas deben filtrar por año y mes, pero seguir almacenando y analizando el periodo contable en formato `yyyymm`.

11. `ADM_Persona` debe aislarse por `IdEmpresa`, manteniendo documento unico por empresa.
12. El registro de personas debe permitir ubigeo en cascada y activar automaticamente cliente y proveedor segun los checks del formulario.
13. Las ayudas de clientes y proveedores en ventas y compras deben resolver busqueda incremental con seleccion por texto y conservar el Id oculto para grabacion.
14. Las ayudas de compras y ventas deben permitir alta rapida; el registro minimo debe grabar en `ADM_Persona` y activar automaticamente `ADM_Cliente` o `ADM_Proveedor`, dejando `CodigoUbigeo = 150101` cuando el flujo operativo no lo solicite.

## Siguiente tramo recomendado

1. Ejecutar los SP y scripts nuevos en `Dbsisadm` para habilitar paginacion, tipos de comprobante y configuracion contable extendida.
2. Validar visualmente todos los listados web con filtros, paginacion y apertura de edicion desde la grilla.
3. Continuar con el mantenimiento de parametros contables complementarios.
4. Construir reportes contables sobre la misma logica de empresa activa y periodo.
5. Luego pasar a cierres mensuales y controles de consistencia contable.
