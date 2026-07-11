# Propuesta inicial de seguridad por cuenta administradora

Fecha: 10/07/2026
Proyecto: SistemaAdministrativoWeb

## 1. Objetivo

Implementar un modelo de seguridad que permita:

- autenticar usuarios con ASP.NET Identity;
- vincular cada usuario a una cuenta administradora;
- asignar una o varias empresas de esa cuenta a cada usuario;
- controlar permisos por opcion del sistema;
- separar opciones de alcance cuenta y opciones de alcance empresa.

## 2. Regla base del modelo

Todo usuario operativo del sistema debe estar ligado primero a una cuenta administradora.

Luego se define:

- a que empresas puede ingresar;
- que rol tiene dentro de la cuenta;
- que permisos tiene sobre cada opcion.

## 3. Estructura funcional propuesta

### 3.1 Niveles

Se manejaran tres niveles de seguridad:

1. Identity
   Autentica al usuario.

2. Cuenta administradora
   Determina a que organizacion pertenece el usuario.

3. Empresa y modulo
   Determina en que empresas puede operar y que acciones puede realizar.

### 3.2 Alcances de modulo

Los modulos del sistema deben clasificarse por alcance:

- `CUENTA`
- `EMPRESA`

#### Modulos de alcance cuenta

- Dashboard
- Empresas
- Usuarios
- Configuracion
- Mi suscripcion
- Ayuda

#### Modulos de alcance empresa

- Plan de cuentas
- Centros de costo
- Cuentas corrientes
- Personas
- Tipos de cambio
- Origenes
- Cuentas destino
- Configuracion contable
- Asientos
- Compras
- Ventas
- Caja y Bancos
- Transferencias
- Aplicaciones
- Procesos
- Reportes
- Libros Electronicos

## 4. Comportamiento del login

El login no solo debe autenticar. Tambien debe resolver el contexto del usuario.

Flujo propuesto:

1. Identity valida correo y clave.
2. El sistema obtiene el `AspNetUserId`.
3. Busca si el usuario pertenece a una cuenta administradora.
4. Valida si la cuenta administradora esta activa.
5. Lista empresas habilitadas para el usuario.
6. Carga permisos de cuenta y permisos de empresa.
7. Redirige segun el contexto resultante.

### 4.1 Resultados esperados

- Si es `SuperAdmin`, entra al panel de plataforma.
- Si tiene una sola empresa asignada, entra directo a esa empresa.
- Si tiene varias empresas asignadas, pasa primero por selector de empresa.
- Si pertenece a la cuenta pero no tiene empresas asignadas, solo ve modulos de cuenta segun permisos.
- Si no pertenece a ninguna cuenta administradora, no debe operar funcionalmente el sistema.

## 5. Opcion nueva en General

Dentro de `General` se agregaran estas opciones:

- `Usuarios`
- `Configuracion`

## 6. Pantalla General > Usuarios

Esta opcion administrara usuarios de la cuenta administradora.

### 6.1 Funciones principales

- listar usuarios vinculados a la cuenta;
- registrar usuario por correo;
- vincular usuario a una o varias empresas;
- asignar rol de cuenta;
- asignar permisos por modulo;
- activar o desactivar acceso;
- reenviar invitacion o recuperacion de acceso.

### 6.2 Regla de negocio

Un usuario puede pertenecer a una sola cuenta administradora activa.

Dentro de esa cuenta puede:

- no tener empresas aun;
- tener una empresa;
- tener varias empresas.

### 6.3 Perfiles base sugeridos

- `AdministradorCuenta`
- `Supervisor`
- `Operador`
- `Consulta`

Estos perfiles no reemplazan los permisos por opcion. Solo sirven como base inicial.

### 6.4 Registro de usuario

Cuando se cree un usuario desde esta pantalla, el sistema debe:

1. crear o reutilizar el usuario en Identity;
2. vincularlo a la cuenta administradora;
3. asignarle empresas;
4. asignarle rol base;
5. guardar permisos por modulo cuando corresponda.

## 7. Pantalla General > Configuracion

Esta opcion administrara datos de la cuenta administradora y datos de facturacion.

### 7.1 Seccion Datos de la cuenta administradora

Campos propuestos:

- nombre de la cuenta administradora;
- nombre del responsable principal;
- correo administrativo;
- telefono principal;
- empresa predeterminada;
- estado de la cuenta;
- observacion administrativa.

### 7.2 Seccion Datos de facturacion

Campos propuestos:

- tipo de comprobante preferido: boleta o factura;
- tipo de documento: DNI o RUC;
- numero de documento;
- nombres y apellidos, cuando sea boleta;
- razon social, cuando sea factura;
- correo de facturacion;
- telefono de facturacion;
- direccion fiscal;
- ubigeo;
- distrito;
- provincia;
- departamento;
- observacion de facturacion.

### 7.3 Campos descartados por definicion actual

No se incluiran:

- logo opcional;
- zona horaria;
- moneda base.

## 8. Integracion con API de Migo

Se reutilizara la integracion ya existente con `MigoPadronApiClient`.

Comportamiento esperado:

- si el usuario ingresa un DNI, se consulta nombre completo;
- si el usuario ingresa un RUC, se consulta razon social, direccion y ubigeo;
- la pantalla debe poblar automaticamente los campos recuperados;
- si Migo no devuelve datos, se permitira ingreso manual.

## 9. Modelo de datos propuesto

Se propone extender la seguridad de negocio con las siguientes tablas:

- `SEG_ModuloSistema`
- `SEG_RolCuenta`
- `SEG_RolCuentaPermiso`
- `SEG_UsuarioCuentaAdministradora`
- `SEG_UsuarioCuentaPermiso`
- `SEG_UsuarioEmpresa`
- `SEG_UsuarioEmpresaPermiso`

### 9.1 Proposito resumido

- `SEG_ModuloSistema`
  Catalogo de opciones del sistema.

- `SEG_RolCuenta`
  Perfiles base por cuenta administradora.

- `SEG_RolCuentaPermiso`
  Permisos base del rol sobre cada modulo.

- `SEG_UsuarioCuentaAdministradora`
  Vinculo principal usuario-cuenta administradora.

- `SEG_UsuarioCuentaPermiso`
  Overrides del usuario para modulos de alcance cuenta.

- `SEG_UsuarioEmpresa`
  Empresas habilitadas para el usuario.

- `SEG_UsuarioEmpresaPermiso`
  Overrides del usuario para modulos de alcance empresa.

## 10. Regla de autorizacion

La autorizacion debe resolverse segun el alcance del modulo:

- si el modulo es `CUENTA`, validar contra cuenta administradora;
- si el modulo es `EMPRESA`, validar contra empresa activa;
- si el usuario no tiene empresa activa y el modulo es de empresa, denegar acceso;
- si el usuario no tiene permiso efectivo, denegar acceso aunque pertenezca a la cuenta.

## 11. Flujo de alta inicial recomendado

Se recomienda mantener el alta publica simple y completar el resto con asistente.

Flujo:

1. registro de usuario;
2. primer ingreso;
3. asistente de configuracion inicial;
4. creacion de cuenta administradora;
5. creacion de empresa principal;
6. guardado de datos de facturacion;
7. vinculacion del usuario como administrador de cuenta;
8. asignacion de permisos base.

## 12. Fases sugeridas de implementacion

### Fase 1

- crear catalogo de modulos;
- crear relaciones usuario-cuenta y usuario-empresa;
- crear permisos base;
- resolver contexto post-login.

### Fase 2

- implementar `General > Usuarios`;
- asignacion de empresas;
- asignacion de permisos por modulo.

### Fase 3

- implementar `General > Configuracion`;
- integrar consulta de DNI y RUC con Migo;
- guardar datos de facturacion por cuenta administradora.

## 13. Primera decision cerrada para continuar

Queda definido que:

- el usuario se crea para la cuenta administradora;
- despues se le asignan empresas;
- `Configuracion` no llevara logo;
- `Configuracion` no llevara zona horaria;
- `Configuracion` no llevara moneda base.
