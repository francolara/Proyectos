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

### Nucleo contable inicial

- `CON_PlanCuenta`
- `CON_Origen`
- `CON_Asiento`
- `CON_AsientoDetalle`
- `CON_CorrelativoAsiento`
- `CON_CuentaDestinoRegla`
- `CON_CuentaDestinoReglaDetalle`
- `usp_CON_ListarOrigenesActivos`
- `usp_CON_ListarPlanCuentaPorEmpresa`
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
- Clientes
- Proveedores
- Monedas
- Tipo de cambio

### Fase 3. Contabilidad estructural

- Plan de cuentas por empresa
- Origenes contables
- Reglas de cuentas destino por ejercicio para distribuciones tipo 6/79 u otras equivalencias
- Asiento manual
- Validacion de cuadre Debe/Haber
- Numeracion de asientos por empresa, origen y periodo mensual con reinicio en cada mes

### Fase 4. Mantenimientos contables

- Ejercicio contable
- Periodos contables
- Configuracion de cuentas por defecto
- Parametros de compras y ventas para contabilizacion automatica

### Fase 5. Registro de compras

- Provisiones de compras
- Relacion con proveedor
- Moneda y tipo de cambio
- Generacion automatica de asiento por origen `COM`

### Fase 6. Registro de ventas

- Comprobantes de venta
- Relacion con cliente
- Moneda y tipo de cambio
- Generacion automatica de asiento por origen `VEN`

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

## Siguiente tramo recomendado

1. Crear mantenimiento web de cuentas administradoras y alta de empresas dentro de la cuenta.
2. Crear mantenimiento web de plan de cuentas y de reglas de cuentas destino.
3. Crear mantenimiento web de origenes contables.
4. Crear pantalla de registro de asiento manual con cabecera y detalle.
5. Luego construir compras y ventas para que generen asientos automaticamente.
