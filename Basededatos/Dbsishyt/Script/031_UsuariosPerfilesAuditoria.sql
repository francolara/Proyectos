-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Administración de usuarios/perfiles; migración compatible, correo único y permisos detallados.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Completa sellos de concurrencia faltantes de los registros existentes.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.AspNetUsers',N'Apellidos') IS NULL ALTER TABLE dbo.AspNetUsers ADD Apellidos NVARCHAR(150) NOT NULL CONSTRAINT DF_AspNetUsers_Apellidos DEFAULT N'';
GO
IF COL_LENGTH(N'dbo.AspNetUsers',N'DebeCambiarClave') IS NULL ALTER TABLE dbo.AspNetUsers ADD DebeCambiarClave BIT NOT NULL CONSTRAINT DF_AspNetUsers_CambiarClave DEFAULT 0;
GO
IF COL_LENGTH(N'dbo.AspNetUsers',N'FechaCambioClave') IS NULL ALTER TABLE dbo.AspNetUsers ADD FechaCambioClave DATETIME2(3) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetUsers',N'UltimoIngreso') IS NULL ALTER TABLE dbo.AspNetUsers ADD UltimoIngreso DATETIME2(3) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetUsers',N'UsuarioRegistro') IS NULL ALTER TABLE dbo.AspNetUsers ADD UsuarioRegistro NVARCHAR(128) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetUsers',N'UsuarioActualizacion') IS NULL ALTER TABLE dbo.AspNetUsers ADD UsuarioActualizacion NVARCHAR(128) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetRoles',N'FechaRegistro') IS NULL ALTER TABLE dbo.AspNetRoles ADD FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_AspNetRoles_FechaRegistro DEFAULT SYSUTCDATETIME();
GO
IF COL_LENGTH(N'dbo.AspNetRoles',N'FechaActualizacion') IS NULL ALTER TABLE dbo.AspNetRoles ADD FechaActualizacion DATETIME2(3) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetRoles',N'UsuarioRegistro') IS NULL ALTER TABLE dbo.AspNetRoles ADD UsuarioRegistro NVARCHAR(128) NULL;
GO
IF COL_LENGTH(N'dbo.AspNetRoles',N'UsuarioActualizacion') IS NULL ALTER TABLE dbo.AspNetRoles ADD UsuarioActualizacion NVARCHAR(128) NULL;
GO
IF EXISTS(SELECT UPPER(LTRIM(RTRIM(usuario.Email))) FROM dbo.AspNetUsers AS usuario WHERE NULLIF(LTRIM(RTRIM(usuario.Email)),N'') IS NOT NULL GROUP BY UPPER(LTRIM(RTRIM(usuario.Email))) HAVING COUNT_BIG(1)>1)
    THROW 51701,N'Existen correos duplicados. Revise las cuentas antes de aplicar la migración de usuarios.',1;
UPDATE dbo.AspNetUsers SET NormalizedEmail=UPPER(NULLIF(LTRIM(RTRIM(Email)),N'')) WHERE NULLIF(LTRIM(RTRIM(Email)),N'') IS NOT NULL;
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.AspNetUsers') AND name=N'UX_AspNetUsers_EmailUnico')
    CREATE UNIQUE INDEX UX_AspNetUsers_EmailUnico ON dbo.AspNetUsers(NormalizedEmail) WHERE NormalizedEmail IS NOT NULL;
GO
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CLIENTES.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CLIENTES.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TIENDAS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TIENDAS.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'MATERIALES.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'MATERIALES.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'UNIDADES.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'UNIDADES.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'ZONAS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'ZONAS.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TARIFARIOS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TARIFARIOS.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PRESUPUESTOS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PRESUPUESTOS.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'OC_CLIENTE.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'OC_CLIENTE.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'GUIAS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'GUIAS.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CONFIRMACIONES.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CONFIRMACIONES.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'EJECUCION.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'EJECUCION.CREAR',N'Crear registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CLIENTES.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CLIENTES.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TIENDAS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TIENDAS.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'MATERIALES.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'MATERIALES.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'UNIDADES.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'UNIDADES.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'ZONAS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'ZONAS.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TARIFARIOS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TARIFARIOS.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PRESUPUESTOS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PRESUPUESTOS.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'OC_CLIENTE.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'OC_CLIENTE.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'GUIAS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'GUIAS.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CONFIRMACIONES.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CONFIRMACIONES.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'EJECUCION.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'EJECUCION.MODIFICAR',N'Editar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CLIENTES.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CLIENTES.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TIENDAS.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TIENDAS.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'MATERIALES.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'MATERIALES.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'UNIDADES.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'UNIDADES.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'ZONAS.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'ZONAS.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'TARIFARIOS.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'TARIFARIOS.ACTIVAR',N'Activar o desactivar registros');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PRESUPUESTOS.CAMBIAR_ESTADO') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PRESUPUESTOS.CAMBIAR_ESTADO',N'Cambiar estado');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'OC_CLIENTE.CAMBIAR_ESTADO') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'OC_CLIENTE.CAMBIAR_ESTADO',N'Cambiar estado');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'OC_CLIENTE.DESCARGAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'OC_CLIENTE.DESCARGAR',N'Descargar archivos');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'GUIAS.DESCARGAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'GUIAS.DESCARGAR',N'Descargar archivos');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'CONFIRMACIONES.DESCARGAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'CONFIRMACIONES.DESCARGAR',N'Descargar archivos');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PRESUPUESTOS.DESCARGAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PRESUPUESTOS.DESCARGAR',N'Descargar archivos');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PRESUPUESTOS.IMPRIMIR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PRESUPUESTOS.IMPRIMIR',N'Imprimir');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'REPORTES.IMPRIMIR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'REPORTES.IMPRIMIR',N'Imprimir');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.CREAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.CREAR',N'Crear usuarios');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.MODIFICAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.MODIFICAR',N'Editar datos y perfil');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.ACTIVAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.ACTIVAR',N'Activar o desactivar usuarios');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.RESTABLECER') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.RESTABLECER',N'Restablecer contraseña');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.DESBLOQUEAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.DESBLOQUEAR',N'Desbloquear acceso');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'USUARIOS.SESIONES') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'USUARIOS.SESIONES',N'Cerrar sesiones');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PERFILES.CONSULTAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PERFILES.CONSULTAR',N'Consultar perfiles');
IF NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=N'PERFILES.EDITAR') INSERT INTO dbo.Permiso(Codigo,Descripcion) VALUES(N'PERFILES.EDITAR',N'Administrar perfiles de acceso');
IF NOT EXISTS(SELECT 1 FROM sys.extended_properties WHERE class=0 AND name=N'HYT-Seguridad-031')
BEGIN
UPDATE dbo.AspNetUsers SET ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()) WHERE ConcurrencyStamp IS NULL;
UPDATE dbo.AspNetRoles SET ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()) WHERE ConcurrencyStamp IS NULL;
-- Los roles existentes conservan las capacidades equivalentes a EDITAR/CONSULTAR.
INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso)
SELECT DISTINCT asignacion.RoleId,nuevo.Codigo FROM dbo.RolPermiso AS asignacion
INNER JOIN dbo.Permiso AS nuevo ON LEFT(nuevo.Codigo,CHARINDEX(N'.',nuevo.Codigo)-1)=LEFT(asignacion.CodigoPermiso,CHARINDEX(N'.',asignacion.CodigoPermiso)-1)
WHERE ((asignacion.CodigoPermiso LIKE N'%.EDITAR' AND nuevo.Codigo NOT LIKE N'%.CONSULTAR') OR (asignacion.CodigoPermiso LIKE N'%.CONSULTAR' AND (nuevo.Codigo LIKE N'%.DESCARGAR' OR nuevo.Codigo LIKE N'%.IMPRIMIR')))
AND NOT EXISTS(SELECT 1 FROM dbo.RolPermiso AS actual WHERE actual.RoleId=asignacion.RoleId AND actual.CodigoPermiso=nuevo.Codigo);
INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso) SELECT N'HYT-ADMIN',permiso.Codigo FROM dbo.Permiso AS permiso
WHERE NOT EXISTS(SELECT 1 FROM dbo.RolPermiso AS actual WHERE actual.RoleId=N'HYT-ADMIN' AND actual.CodigoPermiso=permiso.Codigo);
IF NOT EXISTS(SELECT 1 FROM dbo.AspNetRoles AS rol WHERE rol.Id=N'HYT-ADMINISTRACION')
    INSERT INTO dbo.AspNetRoles(Id,Name,NormalizedName,ConcurrencyStamp) VALUES(N'HYT-ADMINISTRACION',N'Administración',N'ADMINISTRACIÓN',CONVERT(NVARCHAR(36),NEWID()));
INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso) SELECT N'HYT-ADMINISTRACION',permiso.Codigo FROM dbo.Permiso AS permiso
WHERE (permiso.Codigo LIKE N'%.CONSULTAR' OR permiso.Codigo LIKE N'%.IMPRIMIR' OR permiso.Codigo LIKE N'%.DESCARGAR' OR permiso.Codigo LIKE N'CONFIRMACIONES.%' OR permiso.Codigo LIKE N'FACTURADO.%')
AND permiso.Codigo NOT LIKE N'USUARIOS.%' AND permiso.Codigo NOT LIKE N'PERFILES.%' AND permiso.Codigo NOT LIKE N'AUDITORIA.%'
AND NOT EXISTS(SELECT 1 FROM dbo.RolPermiso AS actual WHERE actual.RoleId=N'HYT-ADMINISTRACION' AND actual.CodigoPermiso=permiso.Codigo);
    EXEC sys.sp_addextendedproperty @name=N'HYT-Seguridad-031',@value=N'Aplicado';
END;
GO
