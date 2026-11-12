-- Firma: FRANCO LARA | Create date: 11/11/2026 | Homologa unidades con el catálogo SUNAT y actualiza el mantenimiento. Exige tienda y la muestra en el maestro de contactos.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Usuarios, contraseñas, perfiles de acceso y auditoría.
-- Paquete incremental. Ejecutar con sqlcmd -b después de las etapas previas.
USE [SistemaWebHYT];
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO
-- Firma: FRANCO LARA | Create date: 10/11/2026 | Catálogo de detracciones y código opcional por material.

-- Fuente: Script/031_UsuariosPerfilesAuditoria.sql
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


-- Fuente: Funciones/fn_HYTPermiso.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_HYTPermiso(@UsuarioId NVARCHAR(128),@Codigo NVARCHAR(80))
RETURNS BIT
AS
BEGIN
    DECLARE @Resultado BIT=0;
    IF EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId=usuario.Id
        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId
        WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND usuario.DebeCambiarClave=0 AND permiso.CodigoPermiso=@Codigo)
        SET @Resultado=1;
    RETURN @Resultado;
END;
GO


-- Fuente: StoreProcedure/spu_Seguridad_ProtegerAdministrador.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguridad_ProtegerAdministrador
    
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            WHERE usuario.Activo=1 AND NULLIF(usuario.PasswordHash,N'') IS NOT NULL
                AND NULLIF(usuario.SecurityStamp,N'') IS NOT NULL
                AND NOT EXISTS
                (
                    SELECT requerido.Codigo
                    FROM (VALUES (N'USUARIOS.CONSULTAR'),(N'USUARIOS.CREAR'),(N'USUARIOS.MODIFICAR'),
                        (N'USUARIOS.ACTIVAR'),(N'USUARIOS.RESTABLECER'),(N'PERFILES.CONSULTAR'),(N'PERFILES.EDITAR')) AS requerido(Codigo)
                    WHERE NOT EXISTS
                    (
                        SELECT 1 FROM dbo.AspNetUserRoles AS asignacion
                        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId
                        WHERE asignacion.UserId=usuario.Id AND permiso.CodigoPermiso=requerido.Codigo
                    )
                )
        )
            THROW 51702,N'[HYT_VALIDACION] Debe quedar al menos un administrador activo con permisos para administrar usuarios y perfiles.',1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Seguridad_Catalogos.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguridad_Catalogos
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 AND dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        SELECT (SELECT rol.Id,rol.Name FROM dbo.AspNetRoles AS rol ORDER BY rol.Name FOR JSON PATH) AS Perfiles;
        SELECT (SELECT permiso.Codigo,permiso.Descripcion FROM dbo.Permiso AS permiso ORDER BY permiso.Codigo FOR JSON PATH) AS Permisos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_ConsultarAcceso.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConsultarAcceso
    @UsuarioId NVARCHAR(128) = NULL,
    @NombreNormalizado NVARCHAR(256) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF (@UsuarioId IS NULL AND @NombreNormalizado IS NULL) OR (@UsuarioId IS NOT NULL AND @NombreNormalizado IS NOT NULL)
            THROW 51100, N'Indique un unico criterio de identidad.', 1;
        DECLARE @Id NVARCHAR(128);
        SELECT @Id = usuario.Id FROM dbo.AspNetUsers AS usuario
        WHERE (@UsuarioId IS NOT NULL AND usuario.Id = @UsuarioId)
            OR (@NombreNormalizado IS NOT NULL AND usuario.NormalizedUserName = @NombreNormalizado);

        SELECT usuario.Id, usuario.UserName, LTRIM(RTRIM(usuario.Nombres+N' '+usuario.Apellidos)) AS Nombres, usuario.PasswordHash,
            COALESCE(usuario.SecurityStamp, N'') AS SecurityStamp, usuario.Activo, usuario.LockoutEnd,usuario.DebeCambiarClave
        FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @Id;
        SELECT rol.Name FROM dbo.AspNetUserRoles AS asignacion
        INNER JOIN dbo.AspNetRoles AS rol ON rol.Id = asignacion.RoleId
        WHERE asignacion.UserId = @Id AND rol.Name IS NOT NULL ORDER BY rol.Name;
        SELECT DISTINCT permiso.CodigoPermiso FROM dbo.AspNetUserRoles AS asignacion
        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
        WHERE asignacion.UserId = @Id ORDER BY permiso.CodigoPermiso;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_FinalizarAcceso.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_FinalizarAcceso
    @UsuarioId NVARCHAR(128),
    @SelloEsperado NVARCHAR(MAX),
    @PasswordValido BIT,
    @HashActualizado NVARCHAR(MAX) = NULL,
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51101, N'El acceso administra su propia transaccion.', 1;
        DECLARE @Activo BIT, @Sello NVARCHAR(MAX), @Bloqueo DATETIMEOFFSET,
            @Intentos INT, @Exito BIT = 0, @Ahora DATETIMEOFFSET = SYSUTCDATETIME();
        BEGIN TRANSACTION;
        SET @Propia = 1;
        SELECT @Activo = usuario.Activo, @Sello = COALESCE(usuario.SecurityStamp, N''),
            @Bloqueo = usuario.LockoutEnd, @Intentos = usuario.AccessFailedCount
        FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK) WHERE usuario.Id = @UsuarioId;
        IF @Activo = 1 AND @Sello = @SelloEsperado AND (@Bloqueo IS NULL OR @Bloqueo <= @Ahora)
        BEGIN
            IF @PasswordValido = 1
            BEGIN
                UPDATE dbo.AspNetUsers SET AccessFailedCount = 0, LockoutEnd = NULL,
                    PasswordHash = COALESCE(@HashActualizado, PasswordHash),UltimoIngreso=SYSUTCDATETIME()
                WHERE Id = @UsuarioId;
                SET @Exito = 1;
            END
            ELSE
            BEGIN
                SET @Intentos = CASE WHEN @Bloqueo IS NOT NULL THEN 1 ELSE @Intentos + 1 END;
                UPDATE dbo.AspNetUsers SET AccessFailedCount = CASE WHEN @Intentos >= 5 THEN 0 ELSE @Intentos END,
                    LockoutEnd = CASE WHEN @Intentos >= 5 THEN DATEADD(MINUTE, 15, @Ahora) ELSE NULL END
                WHERE Id = @UsuarioId;
            END;
        END;
        IF @Activo IS NOT NULL
        BEGIN
            INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId)
            VALUES (@CorrelacionId, N'AspNetUsers', @UsuarioId,
                CASE WHEN @Exito = 1 THEN N'INICIO_SESION' ELSE N'ACCESO_RECHAZADO' END, @UsuarioId);
        END;
        IF @PasswordValido=0 AND @Intentos>=5 AND @Activo=1
            INSERT INTO dbo.EventoAuditoria(CorrelacionId,Entidad,RegistroId,Accion,UsuarioId) VALUES(@CorrelacionId,N'AspNetUsers',@UsuarioId,N'BLOQUEO_AUTOMATICO',@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Exito AS Exito;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_RestablecerAcceso.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RestablecerAcceso
    @ActorId NVARCHAR(128),
    @UsuarioNormalizado NVARCHAR(256),
    @PasswordHash NVARCHAR(MAX),
    @SecurityStamp NVARCHAR(MAX),
    @Motivo NVARCHAR(1000),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51105, N'El restablecimiento administra su propia transaccion.', 1;
        IF NULLIF(LTRIM(RTRIM(@Motivo)), N'') IS NULL OR NULLIF(@PasswordHash, N'') IS NULL OR NULLIF(@SecurityStamp, N'') IS NULL
            THROW 51106, N'Indique hash, sello y motivo.', 1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
            WHERE usuario.Id = @ActorId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'USUARIOS.RESTABLECER'
        )
            THROW 51107, N'Actor sin permiso para restablecer acceso.', 1;
        BEGIN TRANSACTION;
        SET @Propia = 1;
        DECLARE @UsuarioId NVARCHAR(128);
        SELECT @UsuarioId = usuario.Id FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK)
        WHERE usuario.NormalizedUserName = @UsuarioNormalizado;
        IF @UsuarioId IS NULL THROW 51108, N'Usuario no encontrado.', 1;
        UPDATE dbo.AspNetUsers SET PasswordHash = @PasswordHash, SecurityStamp = @SecurityStamp,
            ConcurrencyStamp = CONVERT(NVARCHAR(36), NEWID()), AccessFailedCount = 0, LockoutEnd = NULL,
            DebeCambiarClave = 1, UsuarioActualizacion = @ActorId, FechaActualizacion = SYSUTCDATETIME() WHERE Id = @UsuarioId;
        DECLARE @Detalle NVARCHAR(MAX) = (SELECT @UsuarioId AS UsuarioAfectado, ORIGINAL_LOGIN() AS LoginDespliegue FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @ActorId, @CorrelacionId, N'AspNetUsers', @UsuarioId,
            N'RESTABLECER_ACCESO', NULL, NULL, NULL, @Motivo, @Detalle;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CLIENTES.CREAR' ELSE N'CLIENTES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoSAP.',1;
        DECLARE @CodigoSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoSAP'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'TipoEntidad' AND LEN(dato.value)>20) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: TipoEntidad.',1;
        DECLARE @TipoEntidad NVARCHAR(20) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.TipoEntidad'))),N'');
        IF @TipoEntidad IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: TipoEntidad.',1;
        DECLARE @TipoIdentificacion NVARCHAR(20);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'NumeroIdentificacion' AND LEN(dato.value)>30) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: NumeroIdentificacion.',1;
        DECLARE @NumeroIdentificacion NVARCHAR(30) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroIdentificacion'))),N'');
        IF @NumeroIdentificacion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: NumeroIdentificacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'RazonSocial' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: RazonSocial.',1;
        DECLARE @RazonSocial NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RazonSocial'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombres' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombres.',1;
        DECLARE @Nombres NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombres'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Apellidos' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Apellidos.',1;
        DECLARE @Apellidos NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Apellidos'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'DireccionFiscal' AND LEN(dato.value)>300) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: DireccionFiscal.',1;
        DECLARE @DireccionFiscal NVARCHAR(300) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.DireccionFiscal'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoUbigeo' AND LEN(dato.value)>6) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoUbigeo.',1;
        DECLARE @CodigoUbigeo NVARCHAR(6) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        DECLARE @UrlCargaDocumentos NVARCHAR(2048);
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.CodigoSAP, anterior.TipoEntidad, anterior.TipoIdentificacion, anterior.NumeroIdentificacion, anterior.RazonSocial, anterior.Nombres, anterior.Apellidos, anterior.DireccionFiscal, anterior.CodigoUbigeo, anterior.UrlCargaDocumentos, anterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=anterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Cliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.ClienteId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'CLIENTES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @TipoEntidad NOT IN (N'JURIDICA',N'NATURAL') THROW 51213,N'[HYT_VALIDACION] Tipo de entidad no válido.',1;
        IF (@TipoEntidad=N'JURIDICA' AND @RazonSocial IS NULL) OR (@TipoEntidad=N'NATURAL' AND (@Nombres IS NULL OR @Apellidos IS NULL)) THROW 51214,N'[HYT_VALIDACION] Completa la identificación del cliente.',1;
        SET @TipoIdentificacion=JSON_VALUE(@Anterior,N'$.TipoIdentificacion');
        IF @TipoEntidad=N'JURIDICA' BEGIN SET @Nombres=NULL; SET @Apellidos=NULL; END;
        ELSE SET @RazonSocial=NULL;
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        SET @UrlCargaDocumentos=JSON_VALUE(@Anterior,N'$.UrlCargaDocumentos');
        DECLARE @Enlaces NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Enlaces');
        IF EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Enlaces') AND (@Enlaces IS NULL OR LEFT(LTRIM(@Enlaces),1)<>N'[')
            THROW 51226,N'[HYT_VALIDACION] La lista de enlaces no es válida.',1;
        DECLARE @ListaEnlaces TABLE(Id BIGINT NULL,Descripcion NVARCHAR(MAX),Url NVARCHAR(MAX),Orden INT);
        IF @Enlaces IS NOT NULL
        BEGIN
            IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Enlaces))>100 OR EXISTS(SELECT 1 FROM OPENJSON(@Enlaces) AS jsonEnlace WHERE jsonEnlace.type<>5)
                THROW 51226,N'[HYT_VALIDACION] Revisa la lista de enlaces (máximo cien).',1;
            INSERT INTO @ListaEnlaces(Id,Descripcion,Url,Orden)
            SELECT datos.Id,NULLIF(LTRIM(RTRIM(datos.Descripcion)),N''),NULLIF(LTRIM(RTRIM(datos.Url)),N''),CONVERT(INT,entrada.[key])
            FROM OPENJSON(@Enlaces) AS entrada
            CROSS APPLY OPENJSON(entrada.value) WITH(Id BIGINT '$.Id',Descripcion NVARCHAR(MAX) '$.Descripcion',Url NVARCHAR(MAX) '$.Url') AS datos;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS enlace WHERE enlace.Descripcion IS NULL OR LEN(enlace.Descripcion)>200 OR enlace.Url IS NULL OR LEN(enlace.Url)>2048
                OR (enlace.Url NOT LIKE N'https://%' AND enlace.Url NOT LIKE N'http://%') OR CHARINDEX(NCHAR(10),enlace.Url)>0 OR CHARINDEX(NCHAR(13),enlace.Url)>0)
                THROW 51226,N'[HYT_VALIDACION] Completa descripción y enlace HTTP o HTTPS válidos.',1;
            IF EXISTS(SELECT enlace.Id FROM @ListaEnlaces AS enlace WHERE enlace.Id IS NOT NULL GROUP BY enlace.Id HAVING COUNT_BIG(1)>1)
                THROW 51226,N'[HYT_VALIDACION] Un enlace está repetido.',1;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NOT NULL AND NOT EXISTS
                (SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.Id=entrada.Id AND enlace.ClienteId=@Id))
                THROW 51226,N'[HYT_VALIDACION] Un enlace no pertenece al cliente o ya no existe.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Cliente',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'CLI-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Cliente AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Cliente (Codigo, CodigoSAP, TipoEntidad, TipoIdentificacion, NumeroIdentificacion, RazonSocial, Nombres, Apellidos, DireccionFiscal, CodigoUbigeo, UrlCargaDocumentos, Activo, UsuarioRegistro)
            VALUES (@Codigo, @CodigoSAP, @TipoEntidad, @TipoIdentificacion, @NumeroIdentificacion, @RazonSocial, @Nombres, @Apellidos, @DireccionFiscal, @CodigoUbigeo, @UrlCargaDocumentos, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Cliente SET Codigo=@Codigo, CodigoSAP=@CodigoSAP, TipoEntidad=@TipoEntidad, TipoIdentificacion=@TipoIdentificacion, NumeroIdentificacion=@NumeroIdentificacion, RazonSocial=@RazonSocial, Nombres=@Nombres, Apellidos=@Apellidos, DireccionFiscal=@DireccionFiscal, CodigoUbigeo=@CodigoUbigeo, UrlCargaDocumentos=@UrlCargaDocumentos, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ClienteId=@Id;
        END;
        IF @Enlaces IS NOT NULL
        BEGIN
            DELETE enlace FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=@Id
                AND NOT EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id=enlace.Id);
            UPDATE enlace SET Descripcion=entrada.Descripcion,Url=entrada.Url,Orden=entrada.Orden,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
            FROM dbo.ClienteEnlaceDocumento AS enlace INNER JOIN @ListaEnlaces AS entrada ON entrada.Id=enlace.Id WHERE enlace.ClienteId=@Id;
            INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,UsuarioRegistro)
            SELECT @Id,entrada.Descripcion,entrada.Url,entrada.Orden,@UsuarioId FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NULL;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.CodigoSAP, posterior.TipoEntidad, posterior.TipoIdentificacion, posterior.NumeroIdentificacion, posterior.RazonSocial, posterior.Nombres, posterior.Apellidos, posterior.DireccionFiscal, posterior.CodigoUbigeo, posterior.UrlCargaDocumentos, posterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=posterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Cliente AS posterior WHERE posterior.ClienteId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Cliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTienda_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTienda_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TIENDAS.CREAR' ELSE N'TIENDAS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TIENDAS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Direccion' AND LEN(dato.value)>300) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Direccion.',1;
        DECLARE @Direccion NVARCHAR(300) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Direccion'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoUbigeo' AND LEN(dato.value)>6) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoUbigeo.',1;
        DECLARE @CodigoUbigeo NVARCHAR(6) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CentroSAPFICO' AND LEN(dato.value)>50) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CentroSAPFICO.',1;
        DECLARE @CentroSAPFICO NVARCHAR(50) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CentroSAPFICO'))),N'');
        DECLARE @ZonaId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.ZonaId'));
        IF JSON_VALUE(@Datos,N'$.ZonaId') IS NOT NULL AND @ZonaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ZonaId.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.Codigo, anterior.Nombre, anterior.Direccion, anterior.CodigoUbigeo, anterior.CentroSAPFICO, anterior.ZonaId, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Tienda AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.TiendaId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'TIENDAS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @ZonaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ZonaId')),-1)<>@ZonaId) AND NOT EXISTS (SELECT 1 FROM dbo.Zona AS destino WITH (HOLDLOCK) WHERE destino.Id=@ZonaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ZonaId.',1;
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Tienda',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TIE-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Tienda AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Tienda (ClienteId, Codigo, Nombre, Direccion, CodigoUbigeo, CentroSAPFICO, ZonaId, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @Codigo, @Nombre, @Direccion, @CodigoUbigeo, @CentroSAPFICO, @ZonaId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Tienda SET Codigo=@Codigo, Nombre=@Nombre, Direccion=@Direccion, CodigoUbigeo=@CodigoUbigeo, CentroSAPFICO=@CentroSAPFICO, ZonaId=@ZonaId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE TiendaId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.Codigo, posterior.Nombre, posterior.Direccion, posterior.CodigoUbigeo, posterior.CentroSAPFICO, posterior.ZonaId, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Tienda AS posterior WHERE posterior.TiendaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Tienda',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantContactoCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantContactoCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CLIENTES.CREAR' ELSE N'CLIENTES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @TiendaId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TiendaId'));
        IF JSON_VALUE(@Datos,N'$.TiendaId') IS NOT NULL AND @TiendaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TiendaId.',1;
        IF @TiendaId IS NULL THROW 51205,N'[HYT_VALIDACION] Selecciona una tienda para el contacto.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Telefono' AND LEN(dato.value)>50) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Telefono.',1;
        DECLARE @Telefono NVARCHAR(50) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Telefono'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Cargo' AND LEN(dato.value)>100) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Cargo.',1;
        DECLARE @Cargo NVARCHAR(100) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Cargo'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Correo' AND LEN(dato.value)>256) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Correo.',1;
        DECLARE @Correo NVARCHAR(256) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Correo'))),N'');
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.TiendaId, anterior.Nombre, anterior.Telefono, anterior.Cargo, anterior.Correo, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ContactoCliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.ContactoId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'CLIENTES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @TiendaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TiendaId')),-1)<>@TiendaId) AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS destino WITH (HOLDLOCK) WHERE destino.TiendaId=@TiendaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TiendaId.',1;
        IF @TiendaId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId AND tienda.ClienteId=@ClienteId) THROW 51217,N'[HYT_VALIDACION] La tienda no pertenece al cliente.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ContactoCliente (ClienteId, TiendaId, Nombre, Telefono, Cargo, Correo, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @TiendaId, @Nombre, @Telefono, @Cargo, @Correo, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ContactoCliente SET TiendaId=@TiendaId, Nombre=@Nombre, Telefono=@Telefono, Cargo=@Cargo, Correo=@Correo, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ContactoId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.TiendaId, posterior.Nombre, posterior.Telefono, posterior.Cargo, posterior.Correo, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ContactoCliente AS posterior WHERE posterior.ContactoId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ContactoCliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantProveedorCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedorCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CLIENTES.CREAR' ELSE N'CLIENTES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoSAPProveedor' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoSAPProveedor.',1;
        DECLARE @CodigoSAPProveedor NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoSAPProveedor'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'ContratoMarco' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: ContratoMarco.',1;
        DECLARE @ContratoMarco NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.ContratoMarco'))),N'');
        DECLARE @VigenciaDesde DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaDesde'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaDesde') IS NOT NULL AND @VigenciaDesde IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaDesde.',1;
        DECLARE @VigenciaHasta DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaHasta'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaHasta') IS NOT NULL AND @VigenciaHasta IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaHasta.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.CodigoSAPProveedor, anterior.ContratoMarco, anterior.VigenciaDesde, anterior.VigenciaHasta, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ProveedorCliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'CLIENTES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @VigenciaDesde IS NOT NULL AND @VigenciaHasta<@VigenciaDesde THROW 51219,N'[HYT_VALIDACION] Revisa las fechas de vigencia.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ProveedorCliente (ClienteId, CodigoSAPProveedor, ContratoMarco, VigenciaDesde, VigenciaHasta, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @CodigoSAPProveedor, @ContratoMarco, @VigenciaDesde, @VigenciaHasta, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ProveedorCliente SET CodigoSAPProveedor=@CodigoSAPProveedor, ContratoMarco=@ContratoMarco, VigenciaDesde=@VigenciaDesde, VigenciaHasta=@VigenciaHasta, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.CodigoSAPProveedor, posterior.ContratoMarco, posterior.VigenciaDesde, posterior.VigenciaHasta, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ProveedorCliente AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ProveedorCliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: Script/061_HomologacionUnidadesHYT.sql
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.UnidadMedida',N'CodigoSunat') IS NULL
    ALTER TABLE dbo.UnidadMedida ADD CodigoSunat NVARCHAR(3) NULL;
GO
SET XACT_ABORT ON;
BEGIN TRANSACTION;

DECLARE @Origen TABLE (Codigo NVARCHAR(30) PRIMARY KEY, Descripcion NVARCHAR(150), CodigoSunat NVARCHAR(3), Alternativo NVARCHAR(30));
INSERT INTO @Origen (Codigo,Descripcion,CodigoSunat,Alternativo) VALUES
(N'CTO',N'Ciento',N'CEN',NULL),
(N'Mil',N'Millar',N'MIL',NULL),
(N'UND',N'Unidad',N'NIU',NULL),
(N'M2',N'M2',N'MTK',NULL),
(N'BAL',N'Balde',N'NIU',NULL),
(N'KL',N'Kilo',N'KGM',NULL),
(N'M3',N'Metro Cubico',N'MTQ',NULL),
(N'PAR',N'Pares',N'PR',NULL),
(N'GLO',N'Paquete',N'PK',NULL),
(N'MT',N'Metros',N'MTR',NULL),
(N'CAJ',N'Caja',N'BX',N'CAJA'),
(N'ROL',N'Rollo',N'RO',NULL),
(N'PLA',N'Plancha',N'NIU',NULL),
(N'PIE',N'Pieza',N'NIU',NULL),
(N'LIS',N'Lister',N'NIU',NULL),
(N'BOL',N'Bolsa',N'BG',N'BLS'),
(N'KG',N'Kilogramo',N'KGM',NULL),
(N'k',N'otro',N'NIU',NULL),
(N'GL',N'GALON',N'GLL',NULL),
(N'FCO',N'FRASCO',N'NIU',NULL),
(N'JGO',N'JUEGO',N'KT',NULL),
(N'P2',N'PIE CUADRADO',N'FOT',NULL),
(N'LT',N'LITRO',N'LTR',NULL),
(N'KIT',N'KIT',N'KT',NULL),
(N'PAQ',N'PAQUETES',N'PK',NULL); -- Misma unidad comercial que Paquete en el origen.

DECLARE @Actor NVARCHAR(128);
SELECT TOP (1) @Actor=usuario.Id
FROM dbo.AspNetUsers AS usuario
WHERE usuario.Activo=1 AND EXISTS (
    SELECT 1 FROM dbo.AspNetUserRoles AS rol
    INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
    WHERE rol.UserId=usuario.Id AND permiso.CodigoPermiso=N'UNIDADES.EDITAR')
ORDER BY usuario.Id;

UPDATE unidad SET CodigoSunat=origen.CodigoSunat,
    FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=COALESCE(@Actor,unidad.UsuarioRegistro)
FROM dbo.UnidadMedida AS unidad
INNER JOIN @Origen AS origen ON unidad.Codigo IN (origen.Codigo,origen.Alternativo)
WHERE NULLIF(TRIM(unidad.CodigoSunat),N'') IS NULL;

IF @Actor IS NOT NULL
    INSERT INTO dbo.UnidadMedida (Codigo,Descripcion,CodigoSunat,Activo,UsuarioRegistro)
    SELECT origen.Codigo,origen.Descripcion,origen.CodigoSunat,1,@Actor
    FROM @Origen AS origen
    WHERE NOT EXISTS (SELECT 1 FROM dbo.UnidadMedida AS unidad WITH (UPDLOCK,HOLDLOCK)
        WHERE unidad.Codigo IN (origen.Codigo,origen.Alternativo));

COMMIT TRANSACTION;
GO

-- Fuente: StoreProcedure/spu_MantUnidadMedida_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, m.Codigo AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            COALESCE(m.CodigoSunat,N'') AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.UnidadMedida AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~' OR m.CodigoSunat LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(40),m.Id) LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_MantUnidadMedida_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.Codigo,
            m.Descripcion,
            m.CodigoSunat,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion
        FROM dbo.UnidadMedida AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'UnidadMedida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'UnidadMedida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_MantUnidadMedida_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'UNIDADES.CREAR' ELSE N'UNIDADES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Codigo' AND LEN(dato.value)>30) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Codigo.',1;
        DECLARE @Codigo NVARCHAR(30) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Codigo'))),N'');
        IF @Codigo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Codigo.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @CodigoSunat NVARCHAR(4000)=NULLIF(UPPER(TRIM(JSON_VALUE(@Datos,N'$.CodigoSunat'))),N'');
        IF LEN(@CodigoSunat)>3 OR @CodigoSunat COLLATE Latin1_General_100_BIN2 LIKE N'%[^A-Z0-9]%'
            THROW 51206,N'[HYT_VALIDACION] El código SUNAT admite hasta tres letras o números.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.CodigoSunat, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.UnidadMedida AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'UNIDADES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.UnidadMedida (Codigo, Descripcion, CodigoSunat, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @CodigoSunat, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.UnidadMedida SET Codigo=@Codigo, Descripcion=@Descripcion, CodigoSunat=CASE WHEN EXISTS(SELECT 1 FROM OPENJSON(@Datos) WHERE [key]=N'CodigoSunat') THEN @CodigoSunat ELSE CodigoSunat END, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.CodigoSunat, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.UnidadMedida AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'UnidadMedida',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        IF ERROR_NUMBER() IN(2601,2627) SET @ErrorMessage=N'[HYT_VALIDACION] Ya existe una unidad con esa abreviatura.';
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_MantZona_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantZona_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'ZONAS.CREAR' ELSE N'ZONAS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'ZONAS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Zona AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'ZONAS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Zona',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'ZON-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Zona AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Zona (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Zona SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Zona AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Zona',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantEspecialidad_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantEspecialidad_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'MATERIALES.CREAR' ELSE N'MATERIALES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Especialidad AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'MATERIALES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Especialidad',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'ESP-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Especialidad AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Especialidad (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Especialidad SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Especialidad AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Especialidad',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTipoServicio_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTipoServicio_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'MATERIALES.CREAR' ELSE N'MATERIALES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TipoServicio AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'MATERIALES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:TipoServicio',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TS-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.TipoServicio AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TipoServicio (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TipoServicio SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TipoServicio AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TipoServicio',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: Tablas/DetraccionCatalogoHYT.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.DetraccionCatalogoHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.DetraccionCatalogoHYT
    (
        Codigo NVARCHAR(3) NOT NULL CONSTRAINT PK_DetraccionCatalogoHYT PRIMARY KEY,
        Nombre NVARCHAR(250) NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL CHECK (Porcentaje BETWEEN 0 AND 100),
        Activo BIT NOT NULL
    );
END;
GO


-- Fuente: Script/050_CatalogoDetraccionesHYT.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
INSERT INTO dbo.DetraccionCatalogoHYT(Codigo,Nombre,Porcentaje,Activo)
SELECT origen.Codigo,origen.Nombre,origen.Porcentaje,origen.Activo
FROM (VALUES
        (N'001', N'Azúcar y melaza de caña', 10.0000, 1),
        (N'002', N'Arroz', 10.0000, 1),
        (N'003', N'Alcohol etílico', 10.0000, 1),
        (N'004', N'Recursos hidrobiológicos', 10.0000, 1),
        (N'005', N'Maíz amarillo duro', 10.0000, 1),
        (N'006', N'Caña de azúcar', 10.0000, 1),
        (N'007', N'Madera', 10.0000, 1),
        (N'008', N'Arena y piedra', 10.0000, 1),
        (N'009', N'Residuos, subproductos, desechos, recortes y desperdicios', 15.0000, 1),
        (N'010', N'Bienes gravados con el IGV o renuncia a la exoneración', 1.5000, 1),
        (N'011', N'Intermediación laboral y tercerización', 12.0000, 1),
        (N'012', N'Animales vivos', 10.0000, 1),
        (N'013', N'Carnes y despojos comestibles', 10.0000, 1),
        (N'014', N'Abonos, cueros y pieles de origen animal', 10.0000, 1),
        (N'015', N'Aceite de pescado', 10.0000, 1),
        (N'016', N'Harina, polvo y pellets de pescado, crustáceos, moluscos y demás invertebrados acuáticos', 10.0000, 1),
        (N'017', N'Arrendamiento de bienes muebles', 10.0000, 1),
        (N'018', N'Mantenimiento y reparación de bienes muebles', 12.0000, 1),
        (N'019', N'Movimiento de carga', 10.0000, 1),
        (N'020', N'Otros servicios empresariales', 12.0000, 1),
        (N'021', N'Leche', 10.0000, 1),
        (N'022', N'Comisión mercantil', 12.0000, 1),
        (N'023', N'Fabricación de bienes por encargo', 12.0000, 1),
        (N'024', N'Servicio de transporte de personas', 10.0000, 1),
        (N'025', N'Servicio de transporte de carga', 4.0000, 1),
        (N'026', N'Transporte de pasajeros', 10.0000, 1),
        (N'027', N'Contratos de construcción', 4.0000, 1),
        (N'028', N'Oro gravado con el IGV', 10.0000, 1),
        (N'029', N'Paprika y otros frutos de los géneros capsicum o pimienta', 10.0000, 1),
        (N'030', N'Minerales metálicos no auríferos', 10.0000, 1),
        (N'031', N'Bienes exonerados del IGV', 1.5000, 1),
        (N'032', N'Oro y demás minerales metálicos exonerados del IGV', 10.0000, 1),
        (N'033', N'Demás servicios gravados con el IGV', 12.0000, 1),
        (N'034', N'Minerales no metálicos', 10.0000, 1),
        (N'035', N'Bien inmueble gravado con IGV', 4.0000, 1),
        (N'036', N'Plomo', 15.0000, 1),
        (N'037', N'Ley 30737', 4.0000, 1)
) AS origen(Codigo,Nombre,Porcentaje,Activo)
WHERE NOT EXISTS(SELECT 1 FROM dbo.DetraccionCatalogoHYT AS actual WHERE actual.Codigo=origen.Codigo);
GO


-- Fuente: Script/051_DetraccionMaterialHYT.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ServicioMaterial',N'CodigoDetraccion') IS NULL
    ALTER TABLE dbo.ServicioMaterial ADD CodigoDetraccion NVARCHAR(3) NULL;
GO
IF OBJECT_ID(N'dbo.FK_ServicioMaterial_DetraccionCatalogoHYT',N'F') IS NULL
    ALTER TABLE dbo.ServicioMaterial WITH CHECK ADD CONSTRAINT FK_ServicioMaterial_DetraccionCatalogoHYT
        FOREIGN KEY (CodigoDetraccion) REFERENCES dbo.DetraccionCatalogoHYT(Codigo);
GO


-- Fuente: StoreProcedure/spu_MantServicioMaterial_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantServicioMaterial_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'MATERIALES.CREAR' ELSE N'MATERIALES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @CodigoInterno NVARCHAR(60);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMaterialSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMaterialSAP.',1;
        DECLARE @CodigoMaterialSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMaterialSAP'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>1500) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(1500) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @UnidadMedidaId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.UnidadMedidaId'));
        IF @UnidadMedidaId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: UnidadMedidaId.',1;
        IF JSON_VALUE(@Datos,N'$.UnidadMedidaId') IS NOT NULL AND @UnidadMedidaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: UnidadMedidaId.',1;
        DECLARE @EspecialidadId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.EspecialidadId'));
        IF JSON_VALUE(@Datos,N'$.EspecialidadId') IS NOT NULL AND @EspecialidadId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: EspecialidadId.',1;
        DECLARE @CodigoDetraccion NVARCHAR(4000)=NULLIF(TRIM(JSON_VALUE(@Datos,N'$.CodigoDetraccion')),N'');
        IF @CodigoDetraccion IS NOT NULL AND LEN(@CodigoDetraccion)<>3
            THROW 51206,N'[HYT_VALIDACION] Código de detracción inválido.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.CodigoInterno, anterior.CodigoMaterialSAP, anterior.Descripcion, anterior.UnidadMedidaId, anterior.EspecialidadId, anterior.CodigoDetraccion, anterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=anterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ServicioMaterial AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.MaterialId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'MATERIALES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF NOT EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'CodigoDetraccion')
            SET @CodigoDetraccion=JSON_VALUE(@Anterior,N'$.CodigoDetraccion');
        IF @CodigoDetraccion IS NOT NULL AND @CodigoDetraccion<>COALESCE(JSON_VALUE(@Anterior,N'$.CodigoDetraccion'),N'')
            AND NOT EXISTS (SELECT 1 FROM dbo.DetraccionCatalogoHYT AS catalogo WITH (HOLDLOCK)
                WHERE catalogo.Codigo=@CodigoDetraccion AND catalogo.Activo=1)
            THROW 51212,N'[HYT_VALIDACION] Selecciona un código de detracción activo.',1;
        IF @UnidadMedidaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.UnidadMedidaId')),-1)<>@UnidadMedidaId) AND NOT EXISTS (SELECT 1 FROM dbo.UnidadMedida AS destino WITH (HOLDLOCK) WHERE destino.Id=@UnidadMedidaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: UnidadMedidaId.',1;
        IF @EspecialidadId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.EspecialidadId')),-1)<>@EspecialidadId) AND NOT EXISTS (SELECT 1 FROM dbo.Especialidad AS destino WITH (HOLDLOCK) WHERE destino.Id=@EspecialidadId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: EspecialidadId.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        IF @Id IS NOT NULL SET @CodigoInterno=JSON_VALUE(@Anterior,N'$.CodigoInterno');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:ServicioMaterial',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @CodigoInterno=N'MAT-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.ServicioMaterial AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.CodigoInterno=@CodigoInterno) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ServicioMaterial (CodigoInterno, CodigoMaterialSAP, Descripcion, UnidadMedidaId, EspecialidadId, CodigoDetraccion, Activo, UsuarioRegistro)
            VALUES (@CodigoInterno, @CodigoMaterialSAP, @Descripcion, @UnidadMedidaId, @EspecialidadId, @CodigoDetraccion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ServicioMaterial SET CodigoInterno=@CodigoInterno, CodigoMaterialSAP=@CodigoMaterialSAP, Descripcion=@Descripcion, UnidadMedidaId=@UnidadMedidaId, EspecialidadId=@EspecialidadId, CodigoDetraccion=@CodigoDetraccion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE MaterialId=@Id;
        END;
        DELETE FROM dbo.MaterialZona WHERE MaterialId=@Id;
        INSERT INTO dbo.MaterialZona (MaterialId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        SELECT @Posterior=(SELECT posterior.CodigoInterno, posterior.CodigoMaterialSAP, posterior.Descripcion, posterior.UnidadMedidaId, posterior.EspecialidadId, posterior.CodigoDetraccion, posterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=posterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ServicioMaterial AS posterior WHERE posterior.MaterialId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ServicioMaterial',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifario_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifario_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TARIFARIOS.CREAR' ELSE N'TARIFARIOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Nombre, anterior.ClienteId, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Tarifario AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.TarifarioId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @Id IS NOT NULL AND COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>COALESCE(@ClienteId,-1) AND EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS v WHERE v.TarifarioId=@Id AND v.Estado<>N'BORRADOR') THROW 51218,N'[HYT_VALIDACION] No se puede reasignar el cliente de un tarifario publicado.',1;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Tarifario',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TAR-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Tarifario AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Tarifario (Codigo, Nombre, ClienteId, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Nombre, @ClienteId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Tarifario SET Codigo=@Codigo, Nombre=@Nombre, ClienteId=@ClienteId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE TarifarioId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Nombre, posterior.ClienteId, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Tarifario AS posterior WHERE posterior.TarifarioId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Tarifario',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TARIFARIOS.CREAR' ELSE N'TARIFARIOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @TarifarioId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TarifarioId'));
        IF @TarifarioId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: TarifarioId.',1;
        IF JSON_VALUE(@Datos,N'$.TarifarioId') IS NOT NULL AND @TarifarioId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TarifarioId.',1;
        DECLARE @VigenciaDesde DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaDesde'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaDesde') IS NOT NULL AND @VigenciaDesde IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaDesde.',1;
        DECLARE @VigenciaHasta DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaHasta'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaHasta') IS NOT NULL AND @VigenciaHasta IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaHasta.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMonedaPredeterminada' AND LEN(dato.value)>3) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMonedaPredeterminada.',1;
        DECLARE @CodigoMonedaPredeterminada NVARCHAR(3) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMonedaPredeterminada'))),N'');
        IF @CodigoMonedaPredeterminada IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoMonedaPredeterminada.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoTipoIGVPredeterminado' AND LEN(dato.value)>10) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoTipoIGVPredeterminado.',1;
        DECLARE @CodigoTipoIGVPredeterminado NVARCHAR(10) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoTipoIGVPredeterminado'))),N'');
        IF @CodigoTipoIGVPredeterminado IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoTipoIGVPredeterminado.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.TarifarioId=@TarifarioId) THROW 51208,N'[HYT_VALIDACION] El tarifario no existe.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.TarifarioId, anterior.NumeroVersion, anterior.VigenciaDesde, anterior.VigenciaHasta, anterior.CodigoMonedaPredeterminada, anterior.CodigoTipoIGVPredeterminado, anterior.Estado, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=anterior.VersionId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TarifarioVersion AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.VersionId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TarifarioId'))<>@TarifarioId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @Id IS NOT NULL AND JSON_VALUE(@Anterior,N'$.Estado')<>N'BORRADOR' THROW 51211,N'[HYT_VALIDACION] La versión publicada es inmutable. Crea una nueva versión.',1;
        IF @TarifarioId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TarifarioId')),-1)<>@TarifarioId) AND NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS destino WITH (HOLDLOCK) WHERE destino.TarifarioId=@TarifarioId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TarifarioId.',1;
        IF @VigenciaDesde IS NOT NULL AND @VigenciaHasta<@VigenciaDesde THROW 51219,N'[HYT_VALIDACION] Revisa las fechas de vigencia.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        DECLARE @NumeroVersion INT;
        SELECT @NumeroVersion=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TarifarioVersion (TarifarioId, VigenciaDesde, VigenciaHasta, CodigoMonedaPredeterminada, CodigoTipoIGVPredeterminado, NumeroVersion, UsuarioRegistro)
            VALUES (@TarifarioId, @VigenciaDesde, @VigenciaHasta, @CodigoMonedaPredeterminada, @CodigoTipoIGVPredeterminado, @NumeroVersion, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TarifarioVersion SET VigenciaDesde=@VigenciaDesde, VigenciaHasta=@VigenciaHasta, CodigoMonedaPredeterminada=@CodigoMonedaPredeterminada, CodigoTipoIGVPredeterminado=@CodigoTipoIGVPredeterminado,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@Id;
        END;
        DELETE FROM dbo.TarifarioVersionZona WHERE VersionId=@Id;
        INSERT INTO dbo.TarifarioVersionZona (VersionId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        SELECT @Posterior=(SELECT posterior.TarifarioId, posterior.NumeroVersion, posterior.VigenciaDesde, posterior.VigenciaHasta, posterior.CodigoMonedaPredeterminada, posterior.CodigoTipoIGVPredeterminado, posterior.Estado, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=posterior.VersionId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS posterior WHERE posterior.VersionId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioPartida_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioPartida_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TARIFARIOS.CREAR' ELSE N'TARIFARIOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @VersionId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.VersionId'));
        IF @VersionId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: VersionId.',1;
        IF JSON_VALUE(@Datos,N'$.VersionId') IS NOT NULL AND @VersionId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VersionId.',1;
        DECLARE @MaterialId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.MaterialId'));
        IF @MaterialId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: MaterialId.',1;
        IF JSON_VALUE(@Datos,N'$.MaterialId') IS NOT NULL AND @MaterialId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: MaterialId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Posicion' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Posicion.',1;
        DECLARE @Posicion NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Posicion'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMaterialSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMaterialSAP.',1;
        DECLARE @CodigoMaterialSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMaterialSAP'))),N'');
        DECLARE @TipoServicioId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.TipoServicioId'));
        IF JSON_VALUE(@Datos,N'$.TipoServicioId') IS NOT NULL AND @TipoServicioId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TipoServicioId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'NumeroNecesidad' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: NumeroNecesidad.',1;
        DECLARE @NumeroNecesidad NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroNecesidad'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMoneda' AND LEN(dato.value)>3) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMoneda.',1;
        DECLARE @CodigoMoneda NVARCHAR(3) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMoneda'))),N'');
        IF @CodigoMoneda IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoMoneda.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoTipoIGV' AND LEN(dato.value)>10) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoTipoIGV.',1;
        DECLARE @CodigoTipoIGV NVARCHAR(10) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoTipoIGV'))),N'');
        IF @CodigoTipoIGV IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoTipoIGV.',1;
        DECLARE @PrecioPEN DECIMAL(19,6) = TRY_CONVERT(DECIMAL(19,6),JSON_VALUE(@Datos,N'$.PrecioPEN'));
        IF JSON_VALUE(@Datos,N'$.PrecioPEN') IS NOT NULL AND @PrecioPEN IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: PrecioPEN.',1;
        IF @PrecioPEN<0 THROW 51207,N'[HYT_VALIDACION] El precio no puede ser negativo.',1;
        DECLARE @PrecioUSD DECIMAL(19,6) = TRY_CONVERT(DECIMAL(19,6),JSON_VALUE(@Datos,N'$.PrecioUSD'));
        IF JSON_VALUE(@Datos,N'$.PrecioUSD') IS NOT NULL AND @PrecioUSD IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: PrecioUSD.',1;
        IF @PrecioUSD<0 THROW 51207,N'[HYT_VALIDACION] El precio no puede ser negativo.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.VersionId=@VersionId AND padre.Estado=N'BORRADOR') THROW 51208,N'[HYT_VALIDACION] Solo se pueden editar partidas de una versión borrador.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.VersionId, anterior.MaterialId, anterior.Posicion, anterior.CodigoMaterialSAP, anterior.TipoServicioId, anterior.NumeroNecesidad, anterior.CodigoMoneda, anterior.CodigoTipoIGV, anterior.PrecioPEN, anterior.PrecioUSD, anterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=anterior.PartidaId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TarifarioPartida AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.PartidaId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.VersionId'))<>@VersionId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @VersionId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.VersionId')),-1)<>@VersionId) AND NOT EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS destino WITH (HOLDLOCK) WHERE destino.VersionId=@VersionId AND destino.Estado = N'BORRADOR') THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: VersionId.',1;
        IF @MaterialId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.MaterialId')),-1)<>@MaterialId) AND NOT EXISTS (SELECT 1 FROM dbo.ServicioMaterial AS destino WITH (HOLDLOCK) WHERE destino.MaterialId=@MaterialId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: MaterialId.',1;
        IF @TipoServicioId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TipoServicioId')),-1)<>@TipoServicioId) AND NOT EXISTS (SELECT 1 FROM dbo.TipoServicio AS destino WITH (HOLDLOCK) WHERE destino.Id=@TipoServicioId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TipoServicioId.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TarifarioPartida (VersionId, MaterialId, Posicion, CodigoMaterialSAP, TipoServicioId, NumeroNecesidad, CodigoMoneda, CodigoTipoIGV, PrecioPEN, PrecioUSD, Activo, UsuarioRegistro)
            VALUES (@VersionId, @MaterialId, @Posicion, @CodigoMaterialSAP, @TipoServicioId, @NumeroNecesidad, @CodigoMoneda, @CodigoTipoIGV, @PrecioPEN, @PrecioUSD, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TarifarioPartida SET MaterialId=@MaterialId, Posicion=@Posicion, CodigoMaterialSAP=@CodigoMaterialSAP, TipoServicioId=@TipoServicioId, NumeroNecesidad=@NumeroNecesidad, CodigoMoneda=@CodigoMoneda, CodigoTipoIGV=@CodigoTipoIGV, PrecioPEN=@PrecioPEN, PrecioUSD=@PrecioUSD, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PartidaId=@Id;
        END;
        DELETE FROM dbo.TarifarioPartidaZona WHERE PartidaId=@Id;
        INSERT INTO dbo.TarifarioPartidaZona (PartidaId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        UPDATE dbo.TarifarioVersion SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@VersionId;
        SELECT @Posterior=(SELECT posterior.VersionId, posterior.MaterialId, posterior.Posicion, posterior.CodigoMaterialSAP, posterior.TipoServicioId, posterior.NumeroNecesidad, posterior.CodigoMoneda, posterior.CodigoTipoIGV, posterior.PrecioPEN, posterior.PrecioUSD, posterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=posterior.PartidaId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioPartida AS posterior WHERE posterior.PartidaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioPartida',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MaterialTarifa_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MaterialTarifa_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@MaterialId BIGINT,
    @CodigoTipoIGV NVARCHAR(10),@PrecioPEN DECIMAL(19,6),@PrecioUSD DECIMAL(19,6)=NULL,@Activo BIT=1,
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TARIFARIOS.CREAR' ELSE N'TARIFARIOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF EXISTS(SELECT permiso.Codigo FROM (VALUES(N'MATERIALES.CONSULTAR'),(N'TARIFARIOS.CONSULTAR'),(N'TARIFARIOS.EDITAR')) AS permiso(Codigo)
            WHERE NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
                INNER JOIN dbo.RolPermiso AS acceso ON acceso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND acceso.CodigoPermiso=permiso.Codigo))
            THROW 51200,N'[HYT_PERMISO] No tiene permiso para administrar el tarifario.',1;
        IF @CodigoTipoIGV IS NULL OR @CodigoTipoIGV NOT IN(N'CON_IGV',N'SIN_IGV') OR (@PrecioPEN IS NULL AND @PrecioUSD IS NULL) OR @PrecioPEN<0 OR @PrecioUSD<0 OR @Activo IS NULL
            THROW 51227,N'[HYT_VALIDACION] Revisa el tipo de IGV, estado e ingresa al menos un precio en soles o dólares.',1;
        IF @PrecioPEN<>ROUND(@PrecioPEN,2) OR @PrecioUSD<>ROUND(@PrecioUSD,2)
            THROW 51227,N'[HYT_VALIDACION] Los precios admiten hasta dos decimales.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT,@VersionId BIGINT,@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@EstadoAnterior NVARCHAR(40),@EstadoNuevo NVARCHAR(40),@CodigoSAP NVARCHAR(60);
        SELECT @CodigoSAP=material.CodigoMaterialSAP FROM dbo.ServicioMaterial AS material WITH(HOLDLOCK) WHERE material.MaterialId=@MaterialId;
        IF NOT EXISTS(SELECT 1 FROM dbo.ServicioMaterial AS material WHERE material.MaterialId=@MaterialId)
            THROW 51227,N'[HYT_VALIDACION] El material no existe.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@VersionId=precio.VersionId
            FROM dbo.TarifarioPartida AS precio WITH(UPDLOCK,HOLDLOCK) INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=precio.VersionId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
            WHERE precio.PartidaId=@Id AND precio.MaterialId=@MaterialId AND precio.VersionFila=@VersionOriginal
              AND (tarifa.EsTarifarioMaterial=1 OR version.Estado=N'BORRADOR');
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El tarifario cambió o no permite edición. Vuelve a consultarlo.',1;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            UPDATE dbo.TarifarioPartida SET CodigoTipoIGV=@CodigoTipoIGV,PrecioPEN=@PrecioPEN,PrecioUSD=@PrecioUSD,Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PartidaId=@Id;
        END
        ELSE
        BEGIN
            DECLARE @Bloqueo INT;
            EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:TarifarioMaterialDirecto',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @Bloqueo<0 THROW 51227,N'[HYT_VALIDACION] No se pudo guardar. Intenta nuevamente.',1;
            SELECT @TarifarioId=tarifa.TarifarioId FROM dbo.Tarifario AS tarifa WITH(UPDLOCK,HOLDLOCK) WHERE tarifa.EsTarifarioMaterial=1;
            IF @TarifarioId IS NULL
            BEGIN
                DECLARE @Codigo NVARCHAR(40),@Numero BIGINT;
                WHILE 1=1
                BEGIN
                    SET @Numero=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                    SET @Codigo=N'TAR-MAT-'+CONVERT(NVARCHAR(20),@Numero);
                    IF NOT EXISTS(SELECT 1 FROM dbo.Tarifario AS tarifa WHERE tarifa.Codigo=@Codigo) BREAK;
                END;
                INSERT INTO dbo.Tarifario(Codigo,Nombre,Activo,UsuarioRegistro,EsTarifarioMaterial) VALUES(@Codigo,N'Tarifario de materiales',1,@UsuarioId,1);
                SET @TarifarioId=CONVERT(BIGINT,SCOPE_IDENTITY());
                INSERT INTO dbo.TarifarioVersion(TarifarioId,NumeroVersion,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
                VALUES(@TarifarioId,1,N'PEN',N'SIN_IGV',N'VIGENTE',@UsuarioId);
            END;
            SELECT @VersionId=version.VersionId FROM dbo.TarifarioVersion AS version WHERE version.TarifarioId=@TarifarioId AND version.NumeroVersion=1;
            IF @VersionId IS NULL THROW 51227,N'[HYT_VALIDACION] No se pudo preparar el tarifario.',1;
            INSERT INTO dbo.TarifarioPartida(VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,CodigoMaterialSAP,Activo,UsuarioRegistro)
            VALUES(@VersionId,@MaterialId,N'PEN',@CodigoTipoIGV,@PrecioPEN,@PrecioUSD,@CodigoSAP,@Activo,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.TarifarioPartidaZona(PartidaId,ZonaId) SELECT @Id,zona.ZonaId FROM dbo.MaterialZona AS zona WHERE zona.MaterialId=@MaterialId;
        END;
        SELECT @Nuevo=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioPartida AS precio WHERE precio.PartidaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');SET @EstadoNuevo=JSON_VALUE(@Nuevo,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN @EstadoAnterior<>@EstadoNuevo THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioPartida',@Registro,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH;
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Publicar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Publicar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT, @Numero INT, @Estado NVARCHAR(20);
        SELECT @TarifarioId=v.TarifarioId FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        IF @TarifarioId IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión ya no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS t WITH (UPDLOCK,HOLDLOCK) WHERE t.TarifarioId=@TarifarioId AND t.Activo=1)
            THROW 51222,N'[HYT_VALIDACION] El tarifario no está activo.',1;
        SELECT @Estado=v.Estado FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK)
        WHERE v.VersionId=@Id AND v.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión cambió. Vuelve a consultarla.',1;
        IF @Estado<>N'BORRADOR' THROW 51223,N'[HYT_VALIDACION] Solo se publica una versión borrador.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id AND p.Activo=1)
            THROW 51224,N'[HYT_VALIDACION] Agrega al menos una partida activa antes de publicar.',1;
        IF EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK)
            INNER JOIN dbo.ServicioMaterial AS material WITH (HOLDLOCK) ON material.MaterialId=p.MaterialId
            INNER JOIN dbo.UnidadMedida AS unidad WITH (HOLDLOCK) ON unidad.Id=material.UnidadMedidaId
            WHERE p.VersionId=@Id AND p.Activo=1 AND (p.PrecioPEN IS NULL OR material.Activo=0 OR unidad.Activo=0))
            THROW 51225,N'[HYT_VALIDACION] Revisa precios PEN, materiales y unidades de las partidas activas.',1;
        DECLARE @Anterior NVARCHAR(MAX), @Nuevo NVARCHAR(MAX), @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        SELECT @Anterior=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        UPDATE dbo.TarifarioVersion SET Estado=N'VIGENTE',FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@Id;
        SELECT @Nuevo=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'PUBLICAR',N'BORRADOR',N'VIGENTE',NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Clonar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Clonar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.CREAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT, @Numero INT, @Estado NVARCHAR(20);
        SELECT @TarifarioId=v.TarifarioId FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        IF @TarifarioId IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión ya no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS t WITH (UPDLOCK,HOLDLOCK) WHERE t.TarifarioId=@TarifarioId AND t.Activo=1)
            THROW 51222,N'[HYT_VALIDACION] El tarifario no está activo.',1;
        SELECT @Estado=v.Estado FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK)
        WHERE v.VersionId=@Id AND v.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión cambió. Vuelve a consultarla.',1;
        SELECT @Numero=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
        INSERT INTO dbo.TarifarioVersion (TarifarioId,NumeroVersion,VigenciaDesde,VigenciaHasta,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
        SELECT v.TarifarioId,@Numero,v.VigenciaDesde,v.VigenciaHasta,v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,N'BORRADOR',@UsuarioId
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @NuevaVersion BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY());
        INSERT INTO dbo.TarifarioVersionZona (VersionId,ZonaId)
        SELECT @NuevaVersion,z.ZonaId FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=@Id;
        DECLARE @Mapa TABLE (AnteriorId BIGINT NOT NULL PRIMARY KEY,NuevoId BIGINT NOT NULL);
        -- MERGE de solo inserción permite relacionar el identificador fuente con el generado.
        MERGE dbo.TarifarioPartida AS destino
        USING (SELECT p.PartidaId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo
               FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id) AS fuente ON 1=0
        WHEN NOT MATCHED THEN INSERT (VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,Posicion,CodigoMaterialSAP,TipoServicioId,NumeroNecesidad,Activo,UsuarioRegistro)
        VALUES (@NuevaVersion,fuente.MaterialId,fuente.CodigoMoneda,fuente.CodigoTipoIGV,fuente.PrecioPEN,fuente.PrecioUSD,fuente.Posicion,fuente.CodigoMaterialSAP,fuente.TipoServicioId,fuente.NumeroNecesidad,fuente.Activo,@UsuarioId)
        OUTPUT fuente.PartidaId,inserted.PartidaId INTO @Mapa (AnteriorId,NuevoId);
        INSERT INTO dbo.TarifarioPartidaZona (PartidaId,ZonaId)
        SELECT mapa.NuevoId,zona.ZonaId FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartidaZona AS zona ON zona.PartidaId=mapa.AnteriorId;
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoNuevo,UsuarioId,CambiosJson)
        SELECT @CorrelacionId,N'TarifarioPartida',CONVERT(NVARCHAR(100),mapa.NuevoId),N'COPIAR',CASE p.Activo WHEN 1 THEN N'true' ELSE N'false' END,@UsuarioId,
            (SELECT mapa.AnteriorId AS PartidaOrigen,p.VersionId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=p.PartidaId FOR JSON PATH)) AS Zonas
             FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartida AS p ON p.PartidaId=mapa.NuevoId;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT @Id AS VersionOrigen,@NuevaVersion AS VersionNueva,@Numero AS NumeroVersion,
            JSON_QUERY((SELECT v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,v.VigenciaDesde,v.VigenciaHasta,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=v.VersionId FOR JSON PATH)) AS Zonas
             FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@NuevaVersion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Valores,
            (SELECT COUNT_BIG(1) FROM @Mapa AS mapa) AS PartidasCopiadas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@NuevaVersion);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'COPIAR',NULL,N'BORRADOR',NULL,NULL,@Cambios;
        SET @Registro=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'GENERAR_COPIA',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @NuevaVersion AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Presupuesto_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PresupuestoId BIGINT,@RevisionId BIGINT,
    @VersionOriginal BINARY(8),@VersionPresupuesto BINARY(8),@NuevoEstado NVARCHAR(20),
    @FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.CAMBIAR_ESTADO') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para cambiar el estado.',1;
        IF @FechaEvento IS NULL OR NULLIF(LTRIM(RTRIM(@Responsable)),N'') IS NULL OR @NuevoEstado IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica fecha y responsable de la acción.',1;
        IF @NuevoEstado IN (N'OBSERVADO',N'RECHAZADO',N'CANCELADO') AND NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId AND p.VersionFila=@VersionPresupuesto)
            THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        DECLARE @Estado NVARCHAR(20),@Empresa NVARCHAR(MAX),@Moneda NVARCHAR(3),@Total DECIMAL(19,2);
        SELECT @Estado=r.Estado,@Empresa=r.EmpresaSnapshotJson,@Moneda=r.CodigoMoneda,@Total=r.Total
        FROM dbo.PresupuestoRevision AS r WITH (UPDLOCK,HOLDLOCK)
        WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@PresupuestoId AND r.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió o no pertenece al presupuesto.',1;
        IF EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.NumeroRevision>(SELECT origen.NumeroRevision FROM dbo.PresupuestoRevision AS origen WHERE origen.RevisionId=@RevisionId))
            THROW 51319,N'[HYT_VALIDACION] Solo se cambia el estado de la última revisión.',1;
        IF NOT ((@Estado IN (N'BORRADOR',N'OBSERVADO') AND @NuevoEstado IN (N'ENVIADO',N'CANCELADO'))
            OR (@Estado=N'ENVIADO' AND @NuevoEstado IN (N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'CANCELADO'))
            OR (@Estado IN (N'RECHAZADO',N'VENCIDO',N'APROBADO') AND @NuevoEstado=N'CANCELADO'))
            THROW 51319,N'[HYT_VALIDACION] La transición no está permitida. Para corregir contenido crea otra revisión.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (HOLDLOCK) WHERE oc.PresupuestoId=@PresupuestoId)
            AND NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId
                WHERE oc.PresupuestoId=@PresupuestoId AND oc.RevisionId=@RevisionId AND oc.Estado<>N'CANCELADA' AND p.TipoAtencion=N'EMERGENCIA'
                AND ((@Estado=N'BORRADOR' AND @NuevoEstado=N'ENVIADO') OR (@Estado=N'ENVIADO' AND @NuevoEstado=N'APROBADO')))
            THROW 51320,N'[HYT_VALIDACION] La revisión vinculada a una OC no permite este cambio.',1;
        IF @NuevoEstado=N'ENVIADO'
        BEGIN
            IF @Total IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId)
                OR EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId AND (linea.Total IS NULL OR linea.PrecioUnitario IS NULL OR linea.CodigoMonedaOrigen<>@Moneda))
                THROW 51321,N'[HYT_VALIDACION] Completa las partidas y los precios en la moneda elegida antes de enviar.',1;
            IF ISJSON(@Empresa)<>1 OR @Empresa IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.Ruc'),N'') IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.RazonSocial'),N'') IS NULL
                THROW 51321,N'[HYT_VALIDACION] Configura los datos de HYT y vuelve a guardar el borrador antes de enviar.',1;
        END;
        UPDATE dbo.PresupuestoRevision SET Estado=@NuevoEstado,FechaAprobacion=CASE WHEN @NuevoEstado=N'APROBADO' THEN @FechaEvento ELSE FechaAprobacion END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionId;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@FechaReal DATETIME2(3)=CONVERT(DATETIME2(3),@FechaEvento),
            @MotivoAudit NVARCHAR(1000)=LEFT(@Motivo,1000),@Detalle NVARCHAR(MAX);
        SET @Detalle=(SELECT @Estado AS EstadoAnterior,@NuevoEstado AS EstadoNuevo,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Observacion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'CAMBIAR_ESTADO',@Estado,@NuevoEstado,@FechaReal,@MotivoAudit,@Detalle;
        COMMIT TRANSACTION;
        SELECT p.PresupuestoId,r.RevisionId,p.VersionFila AS VersionPresupuesto,r.VersionFila AS VersionRevision,r.Estado
        FROM dbo.Presupuesto AS p INNER JOIN dbo.PresupuestoRevision AS r ON r.PresupuestoId=p.PresupuestoId AND r.RevisionId=@RevisionId WHERE p.PresupuestoId=@PresupuestoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Presupuesto_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(MAX),
    @PresupuestoId BIGINT=NULL,@RevisionId BIGINT=NULL,@VersionOriginal BINARY(8)=NULL,
    @VersionPresupuesto BINARY(8)=NULL,@NuevaRevision BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @PresupuestoId IS NULL THEN N'PRESUPUESTOS.CREAR' ELSE N'PRESUPUESTOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para guardar presupuestos.',1;
        IF ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{' THROW 51303,N'[HYT_VALIDACION] Datos no válidos.',1;
        IF (@PresupuestoId IS NULL AND (@RevisionId IS NOT NULL OR @NuevaRevision=1)) OR (@PresupuestoId IS NOT NULL AND @RevisionId IS NULL)
            THROW 51304,N'[HYT_VALIDACION] Revisa el presupuesto y revisión.',1;
        DECLARE @ClienteId BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId')),
            @TiendaId BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TiendaId')),
            @Tipo NVARCHAR(20)=JSON_VALUE(@Datos,N'$.TipoAtencion'),
            @Descripcion NVARCHAR(1500)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N''),
            @Emision DATE=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEmision'),23),
            @Moneda NVARCHAR(3)=JSON_VALUE(@Datos,N'$.Revision.CodigoMoneda'),
            @TipoIGV NVARCHAR(10)=JSON_VALUE(@Datos,N'$.Revision.CodigoTipoIGV'),
            @FechaSolicitud DATE=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.FechaSolicitud'),23);
        IF @ClienteId IS NULL OR @TiendaId IS NULL OR @Descripcion IS NULL OR @Emision IS NULL OR @Tipo IS NULL OR @Tipo NOT IN (N'INVITACION',N'EMERGENCIA')
            OR @Moneda IS NULL OR JSON_VALUE(@Datos,N'$.Revision.CodigoMoneda') NOT IN (N'PEN',N'USD') OR @TipoIGV IS NULL OR @TipoIGV NOT IN (N'CON_IGV',N'SIN_IGV')
            THROW 51305,N'[HYT_VALIDACION] Completa cliente, tienda, tipo, descripción, fecha, moneda y tratamiento IGV.',1;
        IF JSON_VALUE(@Datos,N'$.FechaSolicitud') IS NOT NULL AND @FechaSolicitud IS NULL
            THROW 51305,N'[HYT_VALIDACION] Fecha de solicitud no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>1500)
            THROW 51306,N'[HYT_VALIDACION] La descripción excede 1500 caracteres.',1;
        IF JSON_QUERY(@Datos,N'$.Partidas') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.Partidas')),1)<>N'['
            THROW 51307,N'[HYT_VALIDACION] Indica las partidas, incluso cuando el borrador está vacío.',1;
        IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Datos,N'$.Partidas'))>500
            THROW 51307,N'[HYT_VALIDACION] El presupuesto admite hasta 500 partidas.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') AS elemento WHERE elemento.type<>5)
            THROW 51307,N'[HYT_VALIDACION] Una partida no tiene el formato esperado.',1;
        DECLARE @Campos TABLE (Nombre NVARCHAR(80) NOT NULL,Limite INT NOT NULL);
        INSERT INTO @Campos (Nombre,Limite) VALUES
            (N'NumeroCotizacionCliente',60),(N'Trimestre',40),(N'VigenciaTexto',100),(N'CodigoSAPProveedor',60),(N'ContratoMarco',60),
            (N'Especialidad',150),(N'Coordinador',150),(N'Encargado',150),(N'Firmante',150),(N'CargoFirmante',150),(N'ZonaTexto',200),(N'Consideraciones',20000);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Revision') AS dato INNER JOIN @Campos AS campo ON campo.Nombre=dato.[key] COLLATE DATABASE_DEFAULT WHERE LEN(dato.value)>campo.Limite)
            THROW 51306,N'[HYT_VALIDACION] Un campo de cabecera excede su longitud permitida.',1;
        IF JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion') IS NOT NULL AND TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion'),23) IS NULL
            THROW 51305,N'[HYT_VALIDACION] Fecha de ejecución no válida.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX),@AnteriorRevision NVARCHAR(MAX),@AnteriorPartidas NVARCHAR(MAX),@Estado NVARCHAR(20),
            @MonedaAnterior NVARCHAR(3),@ClienteAnterior BIGINT,@TiendaAnterior BIGINT,@TipoAnterior NVARCHAR(20),@NumeroRevision INT,
            @RevisionOrigen BIGINT=@RevisionId,@Registro NVARCHAR(100),@Accion NVARCHAR(60),@EsNuevo BIT=CASE WHEN @PresupuestoId IS NULL THEN 1 ELSE 0 END;
        IF @EsNuevo=0
        BEGIN
            SELECT @Anterior=(SELECT p.ClienteId,p.TiendaId,p.TipoAtencion,p.Descripcion,p.FechaSolicitud FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),
                @ClienteAnterior=p.ClienteId,@TiendaAnterior=p.TiendaId,@TipoAnterior=p.TipoAtencion
            FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId AND p.VersionFila=@VersionPresupuesto;
            IF @Anterior IS NULL THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
            SELECT @Estado=r.Estado,@MonedaAnterior=r.CodigoMoneda,
                @AnteriorRevision=(SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.NumeroCotizacionCliente,r.ClienteNombre,r.ClienteDocumento,r.ClienteDireccion,r.TiendaNombre,r.TiendaDireccion,
                    r.FechaEjecucion,r.Trimestre,r.VigenciaTexto,r.CodigoSAPProveedor,r.ContratoMarco,r.Especialidad,r.Coordinador,r.Encargado,r.Consideraciones,r.Firmante,r.CargoFirmante,
                    r.CodigoMoneda,r.CodigoTipoIGV,r.FechaDesdeTasaIGV,r.TasaIGVAplicada,r.Subtotal,r.Impuesto,r.Total,r.ZonaTexto,r.EmpresaSnapshotJson,r.DescripcionTrabajo,r.SimboloMoneda,r.TiendaCodigo,r.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.PresupuestoRevision AS r WITH (UPDLOCK,HOLDLOCK)
            WHERE r.PresupuestoId=@PresupuestoId AND r.RevisionId=@RevisionId AND r.VersionFila=@VersionOriginal;
            IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió o no pertenece al presupuesto.',1;
            IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (HOLDLOCK) WHERE oc.PresupuestoId=@PresupuestoId)
                THROW 51310,N'[HYT_VALIDACION] El presupuesto vinculado a una OC no admite cambios ni nuevas revisiones.',1;
            IF @Estado=N'CANCELADO' OR (@NuevaRevision=0 AND @Estado NOT IN (N'BORRADOR',N'OBSERVADO')) OR (@NuevaRevision=1 AND @Estado=N'BORRADOR')
                THROW 51310,N'[HYT_VALIDACION] Edita solo el registrado u observado; para una revisión enviada crea una nueva revisión.',1;
            IF EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.NumeroRevision>(SELECT origen.NumeroRevision FROM dbo.PresupuestoRevision AS origen WHERE origen.RevisionId=@RevisionId))
                THROW 51310,N'[HYT_VALIDACION] Abre la revisión más reciente del presupuesto.',1;
            IF (@ClienteAnterior<>@ClienteId OR @TiendaAnterior<>@TiendaId OR @TipoAnterior<>@Tipo)
                AND (@NuevaRevision=1 OR EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.Estado<>N'BORRADOR') OR EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS g WHERE g.PresupuestoId=@PresupuestoId))
                THROW 51310,N'[HYT_VALIDACION] No se puede cambiar cliente, tienda o tipo cuando ya existe historial enviado o guía.',1;
            IF @MonedaAnterior<>@Moneda AND COALESCE(JSON_VALUE(@Datos,N'$.RevalorizarMoneda'),N'false')<>N'true'
                THROW 51311,N'[HYT_VALIDACION] Confirma la revalorización de todas las partidas al cambiar la moneda.',1;
            SELECT @AnteriorPartidas=(SELECT p.NumeroLinea,p.Grupo,p.MaterialId,p.PartidaTarifarioId,p.Posicion,p.CodigoInterno,p.CodigoMaterialSAP,p.CentroSAPFICO,p.NumeroNecesidad,p.TipoServicio,p.Descripcion,p.Unidad,
                p.Cantidad,p.PrecioOrigen,p.CodigoMonedaOrigen,p.CodigoTipoIGVOrigen,p.PrecioUnitario,p.BaseImponible,p.Impuesto,p.Total,p.ZonasTexto,p.OrigenPrecio,p.MotivoPrecioManual
                FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId=@RevisionId ORDER BY p.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES);
        END;
        IF NOT EXISTS (SELECT 1 FROM dbo.Cliente AS cliente WITH (HOLDLOCK) INNER JOIN dbo.Tienda AS tienda WITH (HOLDLOCK) ON tienda.ClienteId=cliente.ClienteId
            WHERE cliente.ClienteId=@ClienteId AND tienda.TiendaId=@TiendaId AND ((cliente.Activo=1 AND tienda.Activo=1) OR (@EsNuevo=0 AND @ClienteAnterior=@ClienteId AND @TiendaAnterior=@TiendaId)))
            THROW 51312,N'[HYT_VALIDACION] Revisa cliente y tienda; la tienda debe pertenecer al cliente.',1;
        DECLARE @DesdeTasa DATE,@Tasa DECIMAL(9,6);
        IF @EsNuevo=0 AND @NuevaRevision=0 AND CONVERT(DATE,JSON_VALUE(@AnteriorRevision,N'$.FechaEmision'),23)=@Emision
        BEGIN
            SET @DesdeTasa=CONVERT(DATE,JSON_VALUE(@AnteriorRevision,N'$.FechaDesdeTasaIGV'),23);
            SET @Tasa=CONVERT(DECIMAL(9,6),JSON_VALUE(@AnteriorRevision,N'$.TasaIGVAplicada'));
        END
        ELSE
            SELECT @DesdeTasa=tasa.FechaDesde,@Tasa=tasa.Valor FROM dbo.TasaIGV AS tasa WITH (HOLDLOCK) WHERE tasa.FechaDesde<=@Emision AND (tasa.FechaHasta IS NULL OR tasa.FechaHasta>@Emision);
        IF @Tasa IS NULL THROW 51313,N'[HYT_VALIDACION] No existe una tasa IGV configurada para la fecha de emisión.',1;
        CREATE TABLE #Lineas
        (
            NumeroLinea INT NOT NULL PRIMARY KEY,PartidaAnterior BIGINT NULL,Grupo NVARCHAR(200) NULL,MaterialId BIGINT NULL,PartidaTarifarioId BIGINT NULL,
            Posicion NVARCHAR(60) NULL,CodigoInterno NVARCHAR(60) NULL,CodigoMaterialSAP NVARCHAR(60) NULL,CentroSAPFICO NVARCHAR(50) NULL,NumeroNecesidad NVARCHAR(60) NULL,
            TipoServicio NVARCHAR(150) NULL,Descripcion NVARCHAR(1500) NULL,Unidad NVARCHAR(30) NULL,Cantidad DECIMAL(19,6) NOT NULL,PrecioOrigen DECIMAL(19,6) NULL,
            CodigoMonedaOrigen NVARCHAR(3) NULL,CodigoTipoIGVOrigen NVARCHAR(10) NULL,PrecioUnitario DECIMAL(19,6) NULL,BaseImponible DECIMAL(19,2) NULL,Impuesto DECIMAL(19,2) NULL,
            Total DECIMAL(19,2) NULL,ZonasTexto NVARCHAR(1000) NULL,OrigenPrecio NVARCHAR(20) NULL,MotivoPrecioManual NVARCHAR(300) NULL,Conservar BIT NOT NULL DEFAULT 0
        );
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') WITH (Cantidad NVARCHAR(MAX),PrecioManual NVARCHAR(MAX)) AS entrada
            WHERE TRY_CONVERT(DECIMAL(19,6),entrada.Cantidad) IS NULL OR TRY_CONVERT(DECIMAL(19,6),entrada.Cantidad)<=0
            OR entrada.Cantidad LIKE N'%[^0-9.]%' OR (CHARINDEX(N'.',entrada.Cantidad)>0 AND LEN(entrada.Cantidad)-CHARINDEX(N'.',entrada.Cantidad)>2)
            OR (entrada.PrecioManual IS NOT NULL AND (TRY_CONVERT(DECIMAL(19,6),entrada.PrecioManual) IS NULL OR TRY_CONVERT(DECIMAL(19,6),entrada.PrecioManual)<0
                OR entrada.PrecioManual LIKE N'%[^0-9.]%' OR (CHARINDEX(N'.',entrada.PrecioManual)>0 AND LEN(entrada.PrecioManual)-CHARINDEX(N'.',entrada.PrecioManual)>2))))
            THROW 51314,N'[HYT_VALIDACION] Revisa cantidades y precios: positivos, hasta dos decimales y dentro del rango permitido.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') AS item CROSS APPLY OPENJSON(item.value) AS dato
            INNER JOIN (VALUES (N'Grupo',200),(N'Posicion',60),(N'CodigoInterno',60),(N'CodigoMaterialSAP',60),(N'CentroSAPFICO',50),(N'NumeroNecesidad',60),(N'TipoServicio',150),
                (N'Descripcion',1500),(N'Unidad',30),(N'ZonasTexto',1000),(N'MotivoPrecioManual',300)) AS limite(Nombre,Largo) ON limite.Nombre=dato.[key] COLLATE DATABASE_DEFAULT WHERE LEN(dato.value)>limite.Largo)
            THROW 51306,N'[HYT_VALIDACION] Un campo de partida excede su longitud permitida.',1;
        INSERT INTO #Lineas (NumeroLinea,PartidaAnterior,Grupo,MaterialId,PartidaTarifarioId,Posicion,CodigoInterno,CodigoMaterialSAP,CentroSAPFICO,NumeroNecesidad,TipoServicio,Descripcion,Unidad,Cantidad,PrecioOrigen,
            CodigoMonedaOrigen,CodigoTipoIGVOrigen,ZonasTexto,OrigenPrecio,MotivoPrecioManual)
        SELECT CONVERT(INT,item.[key])+1,dato.PartidaId,dato.Grupo,dato.MaterialId,dato.PartidaTarifarioId,dato.Posicion,dato.CodigoInterno,dato.CodigoMaterialSAP,dato.CentroSAPFICO,dato.NumeroNecesidad,dato.TipoServicio,
            NULLIF(LTRIM(RTRIM(dato.Descripcion)),N''),NULLIF(LTRIM(RTRIM(dato.Unidad)),N''),CONVERT(DECIMAL(19,6),dato.Cantidad),TRY_CONVERT(DECIMAL(19,6),dato.PrecioManual),dato.CodigoMonedaOrigen,dato.CodigoTipoIGVOrigen,
            dato.ZonasTexto,dato.OrigenPrecio,NULLIF(LTRIM(RTRIM(dato.MotivoPrecioManual)),N'')
        FROM OPENJSON(@Datos,N'$.Partidas') AS item CROSS APPLY OPENJSON(item.value) WITH
            (PartidaId BIGINT,Grupo NVARCHAR(200),MaterialId BIGINT,PartidaTarifarioId BIGINT,Posicion NVARCHAR(60),CodigoInterno NVARCHAR(60),CodigoMaterialSAP NVARCHAR(60),CentroSAPFICO NVARCHAR(50),NumeroNecesidad NVARCHAR(60),TipoServicio NVARCHAR(150),
            Descripcion NVARCHAR(1500),Unidad NVARCHAR(30),Cantidad NVARCHAR(MAX),PrecioManual NVARCHAR(MAX),CodigoMonedaOrigen NVARCHAR(3),CodigoTipoIGVOrigen NVARCHAR(10),ZonasTexto NVARCHAR(1000),OrigenPrecio NVARCHAR(20),MotivoPrecioManual NVARCHAR(300)) AS dato;
        IF JSON_VALUE(@Datos,N'$.DetalleNeto')=N'true' AND (@TipoIGV<>N'SIN_IGV' OR EXISTS(SELECT 1 FROM #Lineas AS linea WHERE linea.MaterialId IS NULL OR linea.CodigoTipoIGVOrigen<>N'SIN_IGV')) THROW 51317,N'[HYT_VALIDACION] El detalle requiere material y precios sin IGV.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio IS NULL OR linea.OrigenPrecio NOT IN (N'TARIFARIO',N'MANUAL'))
            THROW 51315,N'[HYT_VALIDACION] Selecciona el origen de precio de cada partida.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.PartidaAnterior IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS original WHERE original.PartidaId=linea.PartidaAnterior AND original.RevisionId=@RevisionOrigen))
            THROW 51315,N'[HYT_VALIDACION] Una partida anterior no pertenece a esta revisión.',1;
        IF EXISTS (SELECT linea.PartidaAnterior FROM #Lineas AS linea WHERE linea.PartidaAnterior IS NOT NULL GROUP BY linea.PartidaAnterior HAVING COUNT_BIG(1)>1)
            THROW 51315,N'[HYT_VALIDACION] No se puede repetir el identificador de una partida anterior.',1;
        UPDATE linea SET Conservar=1,MaterialId=original.MaterialId,PrecioOrigen=original.PrecioOrigen,CodigoMonedaOrigen=original.CodigoMonedaOrigen,CodigoTipoIGVOrigen=original.CodigoTipoIGVOrigen,
            Descripcion=COALESCE(linea.Descripcion,original.Descripcion),Unidad=COALESCE(linea.Unidad,original.Unidad),Posicion=COALESCE(linea.Posicion,original.Posicion),CodigoInterno=COALESCE(linea.CodigoInterno,original.CodigoInterno),
            CodigoMaterialSAP=COALESCE(linea.CodigoMaterialSAP,original.CodigoMaterialSAP),ZonasTexto=COALESCE(linea.ZonasTexto,original.ZonasTexto)
        FROM #Lineas AS linea INNER JOIN dbo.PresupuestoPartida AS original ON original.PartidaId=linea.PartidaAnterior AND original.RevisionId=@RevisionOrigen
        WHERE linea.OrigenPrecio=N'TARIFARIO' AND original.OrigenPrecio=N'TARIFARIO' AND linea.PartidaTarifarioId=original.PartidaTarifarioId AND original.CodigoMonedaOrigen=@Moneda AND @ClienteAnterior=@ClienteId;
        UPDATE linea SET MaterialId=material.MaterialId,PrecioOrigen=CASE WHEN @Moneda=N'USD' THEN tarifa.PrecioUSD ELSE tarifa.PrecioPEN END,CodigoMonedaOrigen=@Moneda,CodigoTipoIGVOrigen=tarifa.CodigoTipoIGV,
            Descripcion=COALESCE(linea.Descripcion,material.Descripcion),Unidad=COALESCE(linea.Unidad,unidad.Codigo),Posicion=COALESCE(linea.Posicion,tarifa.Posicion,material.Posicion),
            CodigoInterno=COALESCE(linea.CodigoInterno,material.CodigoInterno),CodigoMaterialSAP=COALESCE(linea.CodigoMaterialSAP,tarifa.CodigoMaterialSAP,material.CodigoMaterialSAP),
            TipoServicio=COALESCE(linea.TipoServicio,servicio.Descripcion),NumeroNecesidad=COALESCE(linea.NumeroNecesidad,tarifa.NumeroNecesidad),
            ZonasTexto=COALESCE(linea.ZonasTexto,(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=tarifa.PartidaId)),Conservar=1
        FROM #Lineas AS linea INNER JOIN dbo.TarifarioPartida AS tarifa WITH (HOLDLOCK) ON tarifa.PartidaId=linea.PartidaTarifarioId AND tarifa.Activo=1
        INNER JOIN dbo.TarifarioVersion AS version WITH (HOLDLOCK) ON version.VersionId=tarifa.VersionId AND version.Estado=N'VIGENTE'
        INNER JOIN dbo.Tarifario AS cabecera WITH (HOLDLOCK) ON cabecera.TarifarioId=version.TarifarioId AND cabecera.Activo=1 AND (cabecera.ClienteId IS NULL OR cabecera.ClienteId=@ClienteId)
        INNER JOIN dbo.ServicioMaterial AS material WITH (HOLDLOCK) ON material.MaterialId=tarifa.MaterialId AND material.Activo=1
        INNER JOIN dbo.UnidadMedida AS unidad WITH (HOLDLOCK) ON unidad.Id=material.UnidadMedidaId AND unidad.Activo=1
        LEFT JOIN dbo.TipoServicio AS servicio ON servicio.Id=tarifa.TipoServicioId
        WHERE linea.OrigenPrecio=N'TARIFARIO' AND linea.Conservar=0 AND (version.VigenciaDesde IS NULL OR version.VigenciaDesde<=@Emision) AND (version.VigenciaHasta IS NULL OR version.VigenciaHasta>=@Emision);
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'TARIFARIO' AND linea.Conservar=0)
            THROW 51316,N'[HYT_VALIDACION] Una tarifa no está disponible para este cliente o fecha. Revisa la selección.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.Descripcion IS NULL OR linea.Unidad IS NULL OR linea.CodigoMonedaOrigen IS NULL OR linea.CodigoMonedaOrigen<>@Moneda
            OR linea.CodigoTipoIGVOrigen IS NULL OR linea.CodigoTipoIGVOrigen NOT IN (N'CON_IGV',N'SIN_IGV') OR (linea.OrigenPrecio=N'MANUAL' AND linea.PrecioOrigen IS NOT NULL AND linea.MotivoPrecioManual IS NULL))
            THROW 51317,N'[HYT_VALIDACION] Completa descripción, unidad, moneda, tratamiento IGV y motivo de precio manual.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'MANUAL' AND linea.MaterialId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.ServicioMaterial AS material WHERE material.MaterialId=linea.MaterialId))
            THROW 51317,N'[HYT_VALIDACION] El material indicado no existe.',1;
        UPDATE linea SET CentroSAPFICO=tienda.CentroSAPFICO,Posicion=COALESCE(material.CodigoInterno,linea.CodigoInterno),CodigoInterno=COALESCE(material.CodigoInterno,linea.CodigoInterno)
        FROM #Lineas AS linea CROSS JOIN dbo.Tienda AS tienda
        LEFT JOIN dbo.ServicioMaterial AS material ON material.MaterialId=linea.MaterialId WHERE tienda.TiendaId=@TiendaId;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'MANUAL' AND linea.PartidaTarifarioId IS NOT NULL AND NOT EXISTS
            (SELECT 1 FROM dbo.TarifarioPartida AS referencia INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=referencia.VersionId
             INNER JOIN dbo.Tarifario AS cabecera ON cabecera.TarifarioId=version.TarifarioId
             WHERE referencia.PartidaId=linea.PartidaTarifarioId AND (linea.MaterialId IS NULL OR referencia.MaterialId=linea.MaterialId) AND (cabecera.ClienteId IS NULL OR cabecera.ClienteId=@ClienteId)))
            THROW 51317,N'[HYT_VALIDACION] La referencia tarifaria no corresponde al material o cliente.',1;
        UPDATE #Lineas SET PrecioUnitario=CONVERT(DECIMAL(19,6),ROUND(
            CASE WHEN @TipoIGV=CodigoTipoIGVOrigen THEN PrecioOrigen WHEN @TipoIGV=N'CON_IGV' THEN PrecioOrigen*(1+@Tasa) ELSE PrecioOrigen/(1+@Tasa) END,2));
        UPDATE #Lineas SET Total=CONVERT(DECIMAL(19,2),ROUND(Cantidad*PrecioUnitario,2));
        UPDATE #Lineas SET BaseImponible=CASE WHEN @TipoIGV=N'CON_IGV' THEN CONVERT(DECIMAL(19,2),ROUND(Total/(1+@Tasa),2)) ELSE Total END;
        UPDATE #Lineas SET Impuesto=CASE WHEN @TipoIGV=N'CON_IGV' THEN Total-BaseImponible ELSE CONVERT(DECIMAL(19,2),ROUND(BaseImponible*@Tasa,2)) END;
        UPDATE #Lineas SET Total=BaseImponible+Impuesto;
        DECLARE @Subtotal DECIMAL(19,2),@Impuesto DECIMAL(19,2),@Total DECIMAL(19,2);
        IF EXISTS (SELECT 1 FROM #Lineas) AND NOT EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.Total IS NULL)
            SELECT @Subtotal=SUM(linea.BaseImponible),@Impuesto=SUM(linea.Impuesto),@Total=SUM(linea.Total) FROM #Lineas AS linea;
        DECLARE @ClienteNombre NVARCHAR(200),@ClienteDocumento NVARCHAR(30),@ClienteDireccion NVARCHAR(600),@TiendaNombre NVARCHAR(200),@TiendaDireccion NVARCHAR(600),@TiendaCodigo NVARCHAR(40),@Empresa NVARCHAR(MAX);
        IF @EsNuevo=0 AND @ClienteAnterior=@ClienteId AND @TiendaAnterior=@TiendaId
            SELECT @ClienteNombre=r.ClienteNombre,@ClienteDocumento=r.ClienteDocumento,@ClienteDireccion=r.ClienteDireccion,@TiendaNombre=r.TiendaNombre,@TiendaDireccion=r.TiendaDireccion,@TiendaCodigo=r.TiendaCodigo,@Empresa=r.EmpresaSnapshotJson
            FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionOrigen;
        ELSE
            SELECT @ClienteNombre=CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END,@ClienteDocumento=cliente.NumeroIdentificacion,
                @ClienteDireccion=CONCAT_WS(N' / ',NULLIF(cliente.DireccionFiscal,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=cliente.CodigoUbigeo)),@TiendaNombre=tienda.Nombre,@TiendaDireccion=CONCAT_WS(N' / ',NULLIF(tienda.Direccion,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=tienda.CodigoUbigeo)),@TiendaCodigo=tienda.Codigo
            FROM dbo.Cliente AS cliente INNER JOIN dbo.Tienda AS tienda ON tienda.ClienteId=cliente.ClienteId AND tienda.TiendaId=@TiendaId WHERE cliente.ClienteId=@ClienteId;
        IF @Empresa IS NULL OR @NuevaRevision=1
            SELECT @Empresa=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        IF @TiendaCodigo IS NULL SELECT @TiendaCodigo=tienda.Codigo FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId;
        IF @EsNuevo=1
        BEGIN
            DECLARE @NumeroInterno NVARCHAR(40);
            EXEC dbo.spu_ObtenerNumeroInterno @TipoRegistro=N'PRESUPUESTO',@NumeroInterno=@NumeroInterno OUTPUT;
            DECLARE @RegistroUtc DATETIME2(3)=SYSUTCDATETIME(),@Periodo CHAR(6),@Secuencia BIGINT,@NumeroPresupuesto NVARCHAR(40);
            SET @Periodo=CONVERT(CHAR(6),CONVERT(DATETIME2,@RegistroUtc AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112);
            SELECT @Secuencia=c.Ultimo FROM dbo.PresupuestoCorrelativoMensual AS c WITH(UPDLOCK,HOLDLOCK) WHERE c.Periodo=@Periodo;
            IF @Secuencia IS NULL BEGIN SET @Secuencia=1;INSERT INTO dbo.PresupuestoCorrelativoMensual(Periodo,Ultimo) VALUES(@Periodo,@Secuencia);END;
            ELSE BEGIN SET @Secuencia=@Secuencia+1;UPDATE dbo.PresupuestoCorrelativoMensual SET Ultimo=@Secuencia WHERE Periodo=@Periodo;END;
            SET @NumeroPresupuesto=CONCAT(@Periodo,CASE WHEN @Secuencia<1000000 THEN RIGHT(CONCAT(N'000000',@Secuencia),6) ELSE CONVERT(NVARCHAR(20),@Secuencia) END);
            INSERT INTO dbo.Presupuesto (NumeroInterno,NumeroPresupuesto,FechaRegistro,ClienteId,TiendaId,TipoAtencion,Descripcion,FechaSolicitud,UsuarioRegistro)
            VALUES (@NumeroInterno,@NumeroPresupuesto,@RegistroUtc,@ClienteId,@TiendaId,@Tipo,@Descripcion,@FechaSolicitud,@UsuarioId);
            SET @PresupuestoId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.Presupuesto SET ClienteId=@ClienteId,TiendaId=@TiendaId,TipoAtencion=@Tipo,Descripcion=@Descripcion,FechaSolicitud=@FechaSolicitud,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        IF @EsNuevo=1 OR @NuevaRevision=1
        BEGIN
            SELECT @NumeroRevision=COALESCE(MAX(r.NumeroRevision),0)+1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId;
            INSERT INTO dbo.PresupuestoRevision (PresupuestoId,NumeroRevision,FechaEmision,ClienteNombre,ClienteDocumento,ClienteDireccion,TiendaNombre,TiendaDireccion,CodigoMoneda,CodigoTipoIGV,FechaDesdeTasaIGV,TasaIGVAplicada,UsuarioRegistro)
            VALUES (@PresupuestoId,@NumeroRevision,@Emision,@ClienteNombre,@ClienteDocumento,@ClienteDireccion,@TiendaNombre,@TiendaDireccion,@Moneda,@TipoIGV,@DesdeTasa,@Tasa,@UsuarioId);
            SET @RevisionId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END;
        UPDATE dbo.PresupuestoRevision SET FechaEmision=@Emision,ClienteNombre=@ClienteNombre,ClienteDocumento=@ClienteDocumento,ClienteDireccion=@ClienteDireccion,TiendaNombre=@TiendaNombre,TiendaDireccion=@TiendaDireccion,TiendaCodigo=@TiendaCodigo,
            NumeroCotizacionCliente=JSON_VALUE(@Datos,N'$.Revision.NumeroCotizacionCliente'),FechaEjecucion=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion'),23),Trimestre=JSON_VALUE(@Datos,N'$.Revision.Trimestre'),
            VigenciaTexto=JSON_VALUE(@Datos,N'$.Revision.VigenciaTexto'),CodigoSAPProveedor=JSON_VALUE(@Datos,N'$.Revision.CodigoSAPProveedor'),ContratoMarco=JSON_VALUE(@Datos,N'$.Revision.ContratoMarco'),Especialidad=JSON_VALUE(@Datos,N'$.Revision.Especialidad'),
            Coordinador=JSON_VALUE(@Datos,N'$.Revision.Coordinador'),Encargado=JSON_VALUE(@Datos,N'$.Revision.Encargado'),Consideraciones=(SELECT dato.value FROM OPENJSON(@Datos,N'$.Revision') AS dato WHERE dato.[key]=N'Consideraciones'),
            Firmante=JSON_VALUE(@Datos,N'$.Revision.Firmante'),CargoFirmante=JSON_VALUE(@Datos,N'$.Revision.CargoFirmante'),ZonaTexto=JSON_VALUE(@Datos,N'$.Revision.ZonaTexto'),EmpresaSnapshotJson=@Empresa,DescripcionTrabajo=@Descripcion,
            CodigoMoneda=@Moneda,SimboloMoneda=(SELECT moneda.Simbolo FROM dbo.Moneda AS moneda WHERE moneda.Codigo=@Moneda),CodigoTipoIGV=@TipoIGV,FechaDesdeTasaIGV=@DesdeTasa,TasaIGVAplicada=@Tasa,Subtotal=@Subtotal,Impuesto=@Impuesto,Total=@Total,
            FechaActualizacion=CASE WHEN @EsNuevo=0 AND @NuevaRevision=0 THEN SYSUTCDATETIME() ELSE NULL END,UsuarioActualizacion=CASE WHEN @EsNuevo=0 AND @NuevaRevision=0 THEN @UsuarioId ELSE NULL END
        WHERE RevisionId=@RevisionId;
        DELETE FROM dbo.PresupuestoPartida WHERE RevisionId=@RevisionId;
        INSERT INTO dbo.PresupuestoPartida (RevisionId,NumeroLinea,Grupo,MaterialId,PartidaTarifarioId,Posicion,CodigoInterno,CodigoMaterialSAP,CentroSAPFICO,NumeroNecesidad,TipoServicio,Descripcion,Unidad,Cantidad,
            PrecioOrigen,CodigoMonedaOrigen,CodigoTipoIGVOrigen,PrecioUnitario,BaseImponible,Impuesto,Total,ZonasTexto,OrigenPrecio,MotivoPrecioManual)
        SELECT @RevisionId,linea.NumeroLinea,linea.Grupo,linea.MaterialId,linea.PartidaTarifarioId,linea.Posicion,linea.CodigoInterno,linea.CodigoMaterialSAP,linea.CentroSAPFICO,linea.NumeroNecesidad,linea.TipoServicio,linea.Descripcion,linea.Unidad,linea.Cantidad,
            linea.PrecioOrigen,linea.CodigoMonedaOrigen,linea.CodigoTipoIGVOrigen,linea.PrecioUnitario,linea.BaseImponible,linea.Impuesto,linea.Total,linea.ZonasTexto,linea.OrigenPrecio,linea.MotivoPrecioManual FROM #Lineas AS linea;
        DECLARE @Posterior NVARCHAR(MAX),@PartidasNuevas NVARCHAR(MAX),@Cambios NVARCHAR(MAX);
        SELECT @Posterior=(SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.NumeroCotizacionCliente,r.ClienteNombre,r.ClienteDocumento,r.ClienteDireccion,r.TiendaNombre,r.TiendaDireccion,r.FechaEjecucion,r.Trimestre,r.VigenciaTexto,
            r.CodigoSAPProveedor,r.ContratoMarco,r.Especialidad,r.Coordinador,r.Encargado,r.Consideraciones,r.Firmante,r.CargoFirmante,r.CodigoMoneda,r.CodigoTipoIGV,r.FechaDesdeTasaIGV,r.TasaIGVAplicada,r.Subtotal,r.Impuesto,r.Total,r.ZonaTexto,r.EmpresaSnapshotJson,r.DescripcionTrabajo,r.SimboloMoneda,r.TiendaCodigo,r.Estado
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionId;
        SELECT @PartidasNuevas=(SELECT p.NumeroLinea,p.Grupo,p.MaterialId,p.PartidaTarifarioId,p.Posicion,p.CodigoInterno,p.CodigoMaterialSAP,p.CentroSAPFICO,p.NumeroNecesidad,p.TipoServicio,p.Descripcion,p.Unidad,p.Cantidad,
            p.PrecioOrigen,p.CodigoMonedaOrigen,p.CodigoTipoIGVOrigen,p.PrecioUnitario,p.BaseImponible,p.Impuesto,p.Total,p.ZonasTexto,p.OrigenPrecio,p.MotivoPrecioManual FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId=@RevisionId ORDER BY p.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES);
        SET @Cambios=(SELECT @RevisionOrigen AS RevisionOrigen,JSON_QUERY(@AnteriorRevision) AS Anterior,JSON_QUERY(@AnteriorPartidas) AS PartidasAnteriores,JSON_QUERY(@Posterior) AS Nuevo,JSON_QUERY(@PartidasNuevas) AS PartidasNuevas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        SET @Registro=CONVERT(NVARCHAR(100),@RevisionId);
        SET @Accion=CASE WHEN @EsNuevo=1 THEN N'CREAR' WHEN @NuevaRevision=1 THEN N'NUEVA_REVISION' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoGuardado NVARCHAR(20)=CASE WHEN @EsNuevo=1 OR @NuevaRevision=1 THEN N'BORRADOR' ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,@Accion,@Estado,@EstadoGuardado,NULL,NULL,@Cambios;
        SET @Registro=CONVERT(NVARCHAR(100),@PresupuestoId);
        SET @Cambios=(SELECT JSON_QUERY(@Anterior) AS Anterior,@ClienteId AS ClienteId,@TiendaId AS TiendaId,@Tipo AS TipoAtencion,@Descripcion AS Descripcion,@FechaSolicitud AS FechaSolicitud,@RevisionId AS RevisionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Presupuesto',@Registro,@Accion,NULL,NULL,NULL,NULL,@Cambios;
        DROP TABLE #Lineas;
        COMMIT TRANSACTION;
        SELECT p.PresupuestoId,r.RevisionId,p.NumeroPresupuesto AS NumeroInterno,r.NumeroRevision,p.VersionFila AS VersionPresupuesto,r.VersionFila AS VersionRevision
        FROM dbo.Presupuesto AS p INNER JOIN dbo.PresupuestoRevision AS r ON r.PresupuestoId=p.PresupuestoId AND r.RevisionId=@RevisionId WHERE p.PresupuestoId=@PresupuestoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        IF OBJECT_ID(N'tempdb..#Lineas') IS NOT NULL DROP TABLE #Lineas;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Presupuesto_ObtenerImpresion.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_ObtenerImpresion
    @UsuarioId NVARCHAR(128),@ImpresionId BIGINT,@Descargar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Descargar=1 AND dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.DESCARGAR')=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para descargar el presupuesto.',1;
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar impresiones.',1;
        SELECT impresion.ImpresionId,impresion.RevisionId,revision.PresupuestoId,impresion.Formato,impresion.VersionFormato,impresion.FechaRegistro,
            usuario.Nombres AS Usuario,impresion.ContenidoHTML,impresion.SHA256
        FROM dbo.ImpresionPresupuesto AS impresion INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=impresion.RevisionId
        INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=impresion.UsuarioRegistro WHERE impresion.ImpresionId=@ImpresionId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Presupuesto_RegistrarImpresion.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_RegistrarImpresion
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@RevisionId BIGINT,@VersionOriginal BINARY(8),
    @Formato NVARCHAR(10),@VersionFormato NVARCHAR(20),@ContenidoHTML NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.IMPRIMIR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para generar impresiones.',1;
        IF @Formato IS NULL OR @Formato NOT IN (N'ANTERIOR',N'NUEVO') OR NULLIF(@VersionFormato,N'') IS NULL OR @ContenidoHTML IS NULL
            OR LEN(@ContenidoHTML)<100 OR LEN(@ContenidoHTML)>2000000 OR LOWER(LEFT(@ContenidoHTML,15))<>N'<!doctype html>'
            THROW 51322,N'[HYT_VALIDACION] Formato o contenido de impresión no válido.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Estado NVARCHAR(20);
        SELECT @Estado=revision.Estado FROM dbo.PresupuestoRevision AS revision WITH (UPDLOCK,HOLDLOCK) WHERE revision.RevisionId=@RevisionId AND revision.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió mientras se generaba la impresión. Vuelve a consultarla.',1;
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@ContenidoHTML)),2);
        INSERT INTO dbo.ImpresionPresupuesto (RevisionId,Formato,VersionFormato,ContenidoHTML,SHA256,UsuarioRegistro)
        VALUES (@RevisionId,@Formato,@VersionFormato,@ContenidoHTML,@Huella,@UsuarioId);
        DECLARE @ImpresionId BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY()),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@Cambios NVARCHAR(MAX);
        SET @Cambios=(SELECT @ImpresionId AS ImpresionId,@Formato AS Formato,@VersionFormato AS VersionFormato,@Huella AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'GENERAR_IMPRESION',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @ImpresionId AS ImpresionId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIRMACIONES.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'CONFIRMACION_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId),@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió. Vuelve a consultarla.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END,@Anterior=(SELECT documento.FechaConfirmacion,documento.NumeroConfirmacion,documento.RepresentanteCliente,documento.Observacion,documento.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ConfirmacionFacturacion AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.ConfirmacionId=@Id AND documento.OrdenCompraId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'VIGENTE' AND @NuevoEstado=N'RETIRADA')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        UPDATE dbo.ConfirmacionFacturacion SET Vigente=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@PadreId;
        SELECT @Nuevo=(SELECT actual.FechaConfirmacion,actual.NumeroConfirmacion,actual.RepresentanteCliente,actual.Observacion,actual.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfirmacionFacturacion AS actual WHERE actual.ConfirmacionId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfirmacionFacturacion',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CONFIRMACIONES.CREAR' ELSE N'CONFIRMACIONES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaConfirmacion.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaConfirmacion'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaConfirmacion.',1;
        DECLARE @FechaConfirmacion DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N''));
        IF @FechaConfirmacion IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaConfirmacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroConfirmacion' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroConfirmacion.',1;
        DECLARE @NumeroConfirmacion NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroConfirmacion'))),N'');
        IF @NumeroConfirmacion IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroConfirmacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'RepresentanteCliente' AND LEN(campo.value)>150) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: RepresentanteCliente.',1;
        DECLARE @RepresentanteCliente NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RepresentanteCliente'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'CONFIRMACION_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'CONFIRMACION_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'VIGENTE',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        IF @PresupuestoId IS NULL THROW 51410,N'[HYT_VALIDACION] La OC del cliente no existe.',1;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre AND oc.Estado<>N'CANCELADA')
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió o está cancelada. Vuelve a consultarla.',1;
        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.FechaConfirmacion, documento.NumeroConfirmacion, documento.RepresentanteCliente, documento.Observacion, documento.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END
            FROM dbo.ConfirmacionFacturacion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.ConfirmacionId=@Id AND documento.OrdenCompraId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            AND (@Id IS NULL
                OR (JSON_QUERY(@Datos,N'$.Adjunto') IS NULL
                    AND NOT EXISTS (SELECT 1 FROM OPENJSON(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'))))
                OR @FechaConfirmacion IS DISTINCT FROM TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaConfirmacion'))
                OR @NumeroConfirmacion IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.NumeroConfirmacion')
                OR @RepresentanteCliente IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.RepresentanteCliente')
                OR @Observacion IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.Observacion'))
            THROW 51412,N'[HYT_VALIDACION] En un expediente facturado solo puedes agregar o retirar adjuntos. Retira el marcado facturado para cambiar otros datos.',1;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.ConfirmacionFacturacion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@PadreId AND documento.Vigente=1) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @Estado=N'RETIRADA' THROW 51415,N'[HYT_VALIDACION] La confirmación retirada se conserva como historial. Registra una nueva.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ConfirmacionFacturacion (FechaConfirmacion,NumeroConfirmacion,RepresentanteCliente,Observacion,OrdenCompraId,UsuarioRegistro) VALUES (@FechaConfirmacion,@NumeroConfirmacion,@RepresentanteCliente,@Observacion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.ConfirmacionFacturacion SET FechaConfirmacion=@FechaConfirmacion,NumeroConfirmacion=@NumeroConfirmacion,RepresentanteCliente=@RepresentanteCliente,Observacion=@Observacion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo WHERE vinculo.ConfirmacionId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.ConfirmacionArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.ConfirmacionId=@Id
        WHERE archivo.Retirado=0 AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE TRY_CONVERT(BIGINT,entrada.value)=archivo.ArchivoId);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoAnterior,EstadoNuevo,UsuarioId,Motivo,CambiosJson)
        SELECT @CorrelacionId,N'ArchivoAdjunto',CONVERT(NVARCHAR(100),retirado.ArchivoId),N'RETIRAR',N'VIGENTE',N'RETIRADO',@UsuarioId,@Motivo,
            (SELECT retirado.NombreOriginal AS NombreOriginal FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM @ArchivosRetirados AS retirado;
        IF @ArchivoJson IS NOT NULL
        BEGIN
            DECLARE @Nombre NVARCHAR(255)=JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'),@Clave NVARCHAR(500)=JSON_VALUE(@ArchivoJson,N'$.ClaveAlmacenamiento'),@TipoContenido NVARCHAR(150)=JSON_VALUE(@ArchivoJson,N'$.TipoContenido'),
                @Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@ArchivoJson,N'$.TamanoBytes')),@SHA CHAR(64)=JSON_VALUE(@ArchivoJson,N'$.SHA256'),@Extension NVARCHAR(5);
            SET @Extension=CASE @TipoContenido WHEN N'application/pdf' THEN N'.pdf' WHEN N'image/png' THEN N'.png' WHEN N'image/jpeg' THEN N'.jpg' END;
            IF NULLIF(@Nombre,N'') IS NULL OR @Tamano IS NULL OR @Tamano<=0 OR @Tamano>15728640 OR @Extension IS NULL OR LEN(JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'))>255
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'ConfirmacionCliente/'+dbo.fn_ArchivoSegmento(@NumeroConfirmacion)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.ConfirmacionArchivo (ConfirmacionId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.FechaConfirmacion, actual.NumeroConfirmacion, actual.RepresentanteCliente, actual.Observacion, actual.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfirmacionFacturacion AS actual WHERE actual.ConfirmacionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfirmacionFacturacion',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        DECLARE @EstadoOCAnterior NVARCHAR(40);
        SELECT @EstadoOCAnterior=oc.Estado FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        UPDATE dbo.OrdenCompraCliente
        SET Estado=CASE WHEN Estado=N'REGISTRADA' THEN N'TRABAJO_COMPLETADO' ELSE Estado END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        WHERE OrdenCompraId=@PadreId;
        IF @EstadoOCAnterior=N'REGISTRADA'
        BEGIN
            DECLARE @RegistroOC NVARCHAR(100)=CONVERT(NVARCHAR(100),@PadreId),
                @MotivoOC NVARCHAR(2000)=N'Trabajo completado automáticamente por confirmación del cliente.',
                @CambiosOC NVARCHAR(MAX)=(SELECT @Id AS ConfirmacionId,@NumeroConfirmacion AS NumeroConfirmacion,@RepresentanteCliente AS RepresentanteCliente,
                    JSON_QUERY(N'{"Estado":"REGISTRADA"}') AS Anterior,JSON_QUERY(N'{"Estado":"TRABAJO_COMPLETADO"}') AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@RegistroOC,N'CAMBIAR_ESTADO_AUTOMATICO',@EstadoOCAnterior,N'TRABAJO_COMPLETADO',@FechaConfirmacion,@MotivoOC,@CambiosOC;
        END;
        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'EJECUCION.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'EJECUCION_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=@PadreId,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=documento.Estado,@Anterior=(SELECT documento.Responsable,documento.FechaProgramada,documento.FechaInicioReal,documento.FechaTerminoReal,documento.Descripcion,documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Ejecucion AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.EjecucionId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'PROGRAMADA') OR (@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'PENDIENTE') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'EN_EJECUCION' AND @NuevoEstado=N'TERMINADA') OR (@EstadoAnterior=N'EN_EJECUCION' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'TERMINADA' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'CANCELADA' AND @NuevoEstado=N'PENDIENTE')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        IF @NuevoEstado=N'PROGRAMADA' AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            THROW 51433,N'[HYT_VALIDACION] Completa fecha programada y responsable del trabajo en la ficha.',1;
        IF @NuevoEstado=N'EN_EJECUCION' AND (NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            OR (@TipoAtencion=N'INVITACION' AND (NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado<>N'CANCELADA')
                OR NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL))))
            THROW 51433,N'[HYT_VALIDACION] Completa responsable; invitación requiere OC y programación antes de iniciar.',1;
        IF @NuevoEstado=N'TERMINADA' AND EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND (ejecucion.FechaInicioReal IS NULL OR CONVERT(DATE,ejecucion.FechaInicioReal)>@FechaEvento))
            THROW 51433,N'[HYT_VALIDACION] La fecha de término debe ser igual o posterior al inicio real.',1;
        IF @EstadoAnterior=N'TERMINADA' AND (EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE')))
            THROW 51433,N'[HYT_VALIDACION] Reabre la OC antes de reabrir el trabajo.',1;
        UPDATE dbo.Ejecucion SET Estado=@NuevoEstado,
            FechaInicioReal=CASE WHEN @NuevoEstado=N'EN_EJECUCION' THEN COALESCE(FechaInicioReal,CONVERT(DATETIME2(3),@FechaEvento)) ELSE FechaInicioReal END,
            FechaTerminoReal=CASE WHEN @NuevoEstado=N'TERMINADA' THEN CONVERT(DATETIME2(3),@FechaEvento) WHEN @NuevoEstado=N'EN_EJECUCION' THEN NULL ELSE FechaTerminoReal END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE EjecucionId=@Id;
        SELECT @Nuevo=(SELECT actual.Responsable,actual.FechaProgramada,actual.FechaInicioReal,actual.FechaTerminoReal,actual.Descripcion,actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.Ejecucion AS actual WHERE actual.EjecucionId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Ejecucion',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'EJECUCION.CREAR' ELSE N'EJECUCION.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Responsable' AND LEN(campo.value)>150) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Responsable.',1;
        DECLARE @Responsable NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Responsable'))),N'');
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaProgramada.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaProgramada'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaProgramada.',1;
        DECLARE @FechaProgramada DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N''));
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaInicioReal.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaInicioReal'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaInicioReal.',1;
        DECLARE @FechaInicioReal DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N''));
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaTerminoReal.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaTerminoReal'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaTerminoReal.',1;
        DECLARE @FechaTerminoReal DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N''));
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Descripcion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'EJECUCION_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'EJECUCION_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'PENDIENTE',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SET @PresupuestoId=@PadreId;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de modificar documentos del expediente.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.Responsable, documento.FechaProgramada, documento.FechaInicioReal, documento.FechaTerminoReal, documento.Descripcion, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.Ejecucion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.EjecucionId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.Ejecucion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @FechaTerminoReal IS NOT NULL AND (@FechaInicioReal IS NULL OR @FechaTerminoReal<@FechaInicioReal) THROW 51417,N'[HYT_VALIDACION] Revisa inicio y término reales.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la ejecución antes de editarla.',1;
        IF @Estado=N'PROGRAMADA' AND @FechaProgramada IS NULL THROW 51417,N'[HYT_VALIDACION] La ejecución programada necesita fecha.',1;
        IF @Estado IN (N'EN_EJECUCION',N'TERMINADA') AND (@FechaInicioReal IS NULL OR @Responsable IS NULL OR (@Estado=N'TERMINADA' AND @FechaTerminoReal IS NULL))
            THROW 51417,N'[HYT_VALIDACION] Conserva responsable y fechas reales según el estado.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Ejecucion (Responsable,FechaProgramada,FechaInicioReal,FechaTerminoReal,Descripcion,PresupuestoId,UsuarioRegistro) VALUES (@Responsable,@FechaProgramada,@FechaInicioReal,@FechaTerminoReal,@Descripcion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.Ejecucion SET Responsable=@Responsable,FechaProgramada=@FechaProgramada,FechaInicioReal=@FechaInicioReal,FechaTerminoReal=@FechaTerminoReal,Descripcion=@Descripcion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE EjecucionId=@Id;
        IF JSON_QUERY(@Datos,N'$.Adjunto') IS NOT NULL THROW 51419,N'[HYT_VALIDACION] Esta ficha no admite adjuntos.',1;
        SELECT @Nuevo=(SELECT actual.Responsable, actual.FechaProgramada, actual.FechaInicioReal, actual.FechaTerminoReal, actual.Descripcion, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.Ejecucion AS actual WHERE actual.EjecucionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Ejecucion',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoGUIA_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'GUIAS.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        THROW 51431,N'[HYT_VALIDACION] La guía no utiliza estados. Modifica los datos desde su registro.',1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoGUIA_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'GUIAS.CREAR' ELSE N'GUIAS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaActa.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaActa'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaActa.',1;
        DECLARE @FechaActa DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N''));
        IF @FechaActa IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaActa.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroActa' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroActa.',1;
        DECLARE @NumeroActa NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroActa'))),N'');
        IF @NumeroActa IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroActa.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'GUIA_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'GUIA_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @RevisionAceptadaId BIGINT;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'REGISTRADA',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SET @PresupuestoId=@PadreId;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP (1) @RevisionAceptadaId=r.RevisionId,@EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.FechaActa, documento.NumeroActa, documento.Observacion, documento.FechaAceptacion, documento.RepresentanteCliente, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.GuiaConformidad AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.GuiaId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            AND (@Id IS NULL
                OR (JSON_QUERY(@Datos,N'$.Adjunto') IS NULL
                    AND NOT EXISTS (SELECT 1 FROM OPENJSON(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'))))
                OR @FechaActa IS DISTINCT FROM TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaActa'))
                OR @NumeroActa IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.NumeroActa')
                OR @Observacion IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.Observacion'))
            THROW 51412,N'[HYT_VALIDACION] En un expediente facturado solo puedes agregar o retirar adjuntos. Retira el marcado facturado para cambiar otros datos.',1;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @Id IS NULL AND @EstadoPresupuesto NOT IN (N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO')
            THROW 51416,N'[HYT_VALIDACION] El presupuesto no está vigente para registrar la guía.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.GuiaConformidad (FechaActa,NumeroActa,Observacion,PresupuestoId,UsuarioRegistro) VALUES (@FechaActa,@NumeroActa,@Observacion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.GuiaConformidad SET FechaActa=@FechaActa,NumeroActa=@NumeroActa,Observacion=@Observacion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE GuiaId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.GuiaArchivo AS vinculo WHERE vinculo.GuiaId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.GuiaArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.GuiaId=@Id
        WHERE archivo.Retirado=0 AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE TRY_CONVERT(BIGINT,entrada.value)=archivo.ArchivoId);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoAnterior,EstadoNuevo,UsuarioId,Motivo,CambiosJson)
        SELECT @CorrelacionId,N'ArchivoAdjunto',CONVERT(NVARCHAR(100),retirado.ArchivoId),N'RETIRAR',N'VIGENTE',N'RETIRADO',@UsuarioId,@Motivo,
            (SELECT retirado.NombreOriginal AS NombreOriginal FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM @ArchivosRetirados AS retirado;
        IF @ArchivoJson IS NOT NULL
        BEGIN
            DECLARE @Nombre NVARCHAR(255)=JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'),@Clave NVARCHAR(500)=JSON_VALUE(@ArchivoJson,N'$.ClaveAlmacenamiento'),@TipoContenido NVARCHAR(150)=JSON_VALUE(@ArchivoJson,N'$.TipoContenido'),
                @Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@ArchivoJson,N'$.TamanoBytes')),@SHA CHAR(64)=JSON_VALUE(@ArchivoJson,N'$.SHA256'),@Extension NVARCHAR(5);
            SET @Extension=CASE @TipoContenido WHEN N'application/pdf' THEN N'.pdf' WHEN N'image/png' THEN N'.png' WHEN N'image/jpeg' THEN N'.jpg' END;
            IF NULLIF(@Nombre,N'') IS NULL OR @Tamano IS NULL OR @Tamano<=0 OR @Tamano>15728640 OR @Extension IS NULL OR LEN(JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'))>255
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'Guia/'+dbo.fn_ArchivoSegmento(@NumeroActa)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            IF @TipoContenido<>N'application/pdf' THROW 51419,N'[HYT_VALIDACION] La guía solo admite escaneado PDF.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.GuiaArchivo (GuiaId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.FechaActa, actual.NumeroActa, actual.Observacion, actual.FechaAceptacion, actual.RepresentanteCliente, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.GuiaConformidad AS actual WHERE actual.GuiaId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'GuiaConformidad',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        IF @Anterior IS NULL AND @EstadoPresupuesto<>N'APROBADO'
        BEGIN
            DECLARE @FechaAceptacionAutomatica DATE=CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()));
            UPDATE dbo.PresupuestoRevision SET Estado=N'APROBADO',FechaAprobacion=@FechaAceptacionAutomatica,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionAceptadaId;
            DECLARE @RegistroRevision NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionAceptadaId),
                @MotivoAceptacion NVARCHAR(1000)=N'Aceptación automática por registro de GUIA.',
                @DetalleAceptacion NVARCHAR(MAX)=(SELECT @EstadoPresupuesto AS EstadoAnterior,N'APROBADO' AS EstadoNuevo,
                    @FechaAceptacionAutomatica AS FechaEvento,N'GUIA' AS TipoDocumento,@Id AS DocumentoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@RegistroRevision,N'CAMBIAR_ESTADO',
                @EstadoPresupuesto,N'APROBADO',NULL,@MotivoAceptacion,@DetalleAceptacion;
        END;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoOC_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'OC_CLIENTE.CAMBIAR_ESTADO') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'OC_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=@PadreId,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=documento.Estado,@Anterior=(SELECT documento.RevisionId,documento.NumeroOCCliente,documento.FechaEmisionOC,documento.FechaRecepcionOC,documento.Importe,documento.Observacion,documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.OrdenCompraCliente AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'REGISTRADA' AND @NuevoEstado=N'TRABAJO_COMPLETADO') OR (@EstadoAnterior=N'REGISTRADA' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'CERRADA_OPERATIVAMENTE' AND @NuevoEstado=N'TRABAJO_COMPLETADO') OR (@EstadoAnterior=N'CANCELADA' AND @NuevoEstado=N'REGISTRADA')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        IF @NuevoEstado=N'TRABAJO_COMPLETADO'
            AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')
            AND NOT EXISTS(SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1)
            THROW 51433,N'[HYT_VALIDACION] Registra la confirmación del cliente o la ejecución terminada para completar el trabajo.',1;
        IF @NuevoEstado=N'CERRADA_OPERATIVAMENTE' AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')
            THROW 51433,N'[HYT_VALIDACION] Para cerrar la OC registra la ejecución terminada del trabajo.',1;
        IF @NuevoEstado=N'CERRADA_OPERATIVAMENTE' AND NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId)
            THROW 51433,N'[HYT_VALIDACION] Para cerrar la OC se requiere la guía registrada.',1;
        IF @NuevoEstado=N'CANCELADA' AND (EXISTS(SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1)
            OR EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'EN_EJECUCION'))
            THROW 51433,N'[HYT_VALIDACION] Retira la confirmación, detén la ejecución antes de cancelar la OC.',1;
        UPDATE dbo.OrdenCompraCliente SET Estado=@NuevoEstado,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@Id;
        SELECT @Nuevo=(SELECT actual.RevisionId,actual.NumeroOCCliente,actual.FechaEmisionOC,actual.FechaRecepcionOC,actual.Importe,actual.Observacion,actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.OrdenCompraCliente AS actual WHERE actual.OrdenCompraId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoOC_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'OC_CLIENTE.CREAR' ELSE N'OC_CLIENTE.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N'') IS NOT NULL AND TRY_CONVERT(BIGINT,NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: RevisionId.',1;
        DECLARE @RevisionId BIGINT=TRY_CONVERT(BIGINT,NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N''));
        IF @RevisionId IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: RevisionId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroOCCliente' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroOCCliente.',1;
        DECLARE @NumeroOCCliente NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroOCCliente'))),N'');
        IF @NumeroOCCliente IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroOCCliente.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaEmisionOC.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaEmisionOC'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaEmisionOC.',1;
        DECLARE @FechaEmisionOC DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N''));
        IF @FechaEmisionOC IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaEmisionOC.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaRecepcionOC.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaRecepcionOC.',1;
        DECLARE @FechaRecepcionOC DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N''));
        IF @FechaRecepcionOC IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaRecepcionOC.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N'') IS NOT NULL AND TRY_CONVERT(DECIMAL(19,2),NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: Importe.',1;
        DECLARE @Importe DECIMAL(19,2)=TRY_CONVERT(DECIMAL(19,2),NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N''));
        IF @Importe<0 OR (JSON_VALUE(@Datos,N'$.Importe') IS NOT NULL AND (PATINDEX(N'%[^0-9.]%',JSON_VALUE(@Datos,N'$.Importe'))>0 OR (CHARINDEX(N'.',JSON_VALUE(@Datos,N'$.Importe'))>0 AND LEN(JSON_VALUE(@Datos,N'$.Importe'))-CHARINDEX(N'.',JSON_VALUE(@Datos,N'$.Importe'))>2))) THROW 51408,N'[HYT_VALIDACION] Importe no negativo con hasta dos decimales.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'OC_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'OC_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'REGISTRADA',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SET @PresupuestoId=@PadreId;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.RevisionId, documento.NumeroOCCliente, documento.FechaEmisionOC, documento.FechaRecepcionOC, documento.Importe, documento.Observacion, documento.NumeroRegistroInterno, documento.CodigoMoneda, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.OrdenCompraCliente AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            AND (@Id IS NULL
                OR (JSON_QUERY(@Datos,N'$.Adjunto') IS NULL
                    AND NOT EXISTS (SELECT 1 FROM OPENJSON(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'))))
                OR @RevisionId IS DISTINCT FROM TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.RevisionId'))
                OR @NumeroOCCliente IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.NumeroOCCliente')
                OR @FechaEmisionOC IS DISTINCT FROM TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaEmisionOC'))
                OR @FechaRecepcionOC IS DISTINCT FROM TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaRecepcionOC'))
                OR @Importe IS DISTINCT FROM TRY_CONVERT(DECIMAL(19,2),JSON_VALUE(@Anterior,N'$.Importe'))
                OR @Observacion IS DISTINCT FROM JSON_VALUE(@Anterior,N'$.Observacion'))
            THROW 51412,N'[HYT_VALIDACION] En un expediente facturado solo puedes agregar o retirar adjuntos. Retira el marcado facturado para cambiar otros datos.',1;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        DECLARE @EstadoRevision NVARCHAR(20),@TotalRevision DECIMAL(19,2);
        SELECT @EstadoRevision=r.Estado,@TotalRevision=r.Total,@Moneda=r.CodigoMoneda FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId AND r.RevisionId=@RevisionId;
        IF @EstadoRevision IS NULL OR @TotalRevision IS NULL OR @EstadoRevision NOT IN (N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO')
            THROW 51414,N'[HYT_VALIDACION] El presupuesto debe estar vigente y tener sus importes completos para registrar la OC.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la OC antes de editarla.',1;
        IF @Id IS NOT NULL AND (@Estado=N'CERRADA_OPERATIVAMENTE' OR EXISTS (SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1))
            AND (@RevisionId<>TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.RevisionId')) OR @NumeroOCCliente<>JSON_VALUE(@Anterior,N'$.NumeroOCCliente')
                OR @FechaEmisionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaEmisionOC'),23) OR @FechaRecepcionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaRecepcionOC'),23)
                OR COALESCE(@Importe,-1)<>COALESCE(TRY_CONVERT(DECIMAL(19,2),JSON_VALUE(@Anterior,N'$.Importe')),-1))
            THROW 51415,N'[HYT_VALIDACION] Reabre la OC o retira su confirmación antes de corregir sus datos principales.',1;
        DECLARE @NumeroInterno NVARCHAR(40);
        IF @Id IS NULL EXEC dbo.spu_ObtenerNumeroInterno @TipoRegistro=N'OC_CLIENTE',@NumeroInterno=@NumeroInterno OUTPUT;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.OrdenCompraCliente (RevisionId,NumeroOCCliente,FechaEmisionOC,FechaRecepcionOC,Importe,Observacion,PresupuestoId,UsuarioRegistro,NumeroRegistroInterno,CodigoMoneda) VALUES (@RevisionId,@NumeroOCCliente,@FechaEmisionOC,@FechaRecepcionOC,@Importe,@Observacion,@PadreId,@UsuarioId,@NumeroInterno,@Moneda);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.OrdenCompraCliente SET RevisionId=@RevisionId,NumeroOCCliente=@NumeroOCCliente,FechaEmisionOC=@FechaEmisionOC,FechaRecepcionOC=@FechaRecepcionOC,Importe=@Importe,Observacion=@Observacion,CodigoMoneda=@Moneda,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo WHERE vinculo.OrdenCompraId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.OrdenCompraArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.OrdenCompraId=@Id
        WHERE archivo.Retirado=0 AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE TRY_CONVERT(BIGINT,entrada.value)=archivo.ArchivoId);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoAnterior,EstadoNuevo,UsuarioId,Motivo,CambiosJson)
        SELECT @CorrelacionId,N'ArchivoAdjunto',CONVERT(NVARCHAR(100),retirado.ArchivoId),N'RETIRAR',N'VIGENTE',N'RETIRADO',@UsuarioId,@Motivo,
            (SELECT retirado.NombreOriginal AS NombreOriginal FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM @ArchivosRetirados AS retirado;
        IF @ArchivoJson IS NOT NULL
        BEGIN
            DECLARE @Nombre NVARCHAR(255)=JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'),@Clave NVARCHAR(500)=JSON_VALUE(@ArchivoJson,N'$.ClaveAlmacenamiento'),@TipoContenido NVARCHAR(150)=JSON_VALUE(@ArchivoJson,N'$.TipoContenido'),
                @Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@ArchivoJson,N'$.TamanoBytes')),@SHA CHAR(64)=JSON_VALUE(@ArchivoJson,N'$.SHA256'),@Extension NVARCHAR(5);
            SET @Extension=CASE @TipoContenido WHEN N'application/pdf' THEN N'.pdf' WHEN N'image/png' THEN N'.png' WHEN N'image/jpeg' THEN N'.jpg' END;
            IF NULLIF(@Nombre,N'') IS NULL OR @Tamano IS NULL OR @Tamano<=0 OR @Tamano>15728640 OR @Extension IS NULL OR LEN(JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'))>255
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'OCCliente/'+dbo.fn_ArchivoSegmento(@NumeroOCCliente)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.OrdenCompraArchivo (OrdenCompraId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.RevisionId, actual.NumeroOCCliente, actual.FechaEmisionOC, actual.FechaRecepcionOC, actual.Importe, actual.Observacion, actual.NumeroRegistroInterno, actual.CodigoMoneda, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.OrdenCompraCliente AS actual WHERE actual.OrdenCompraId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        IF @Anterior IS NULL AND @EstadoRevision<>N'APROBADO'
        BEGIN
            DECLARE @FechaAceptacionAutomatica DATE=CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()));
            UPDATE dbo.PresupuestoRevision SET Estado=N'APROBADO',FechaAprobacion=@FechaAceptacionAutomatica,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionId;
            DECLARE @RegistroRevision NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),
                @MotivoAceptacion NVARCHAR(1000)=N'Aceptación automática por registro de OC.',
                @DetalleAceptacion NVARCHAR(MAX)=(SELECT @EstadoRevision AS EstadoAnterior,N'APROBADO' AS EstadoNuevo,
                    @FechaAceptacionAutomatica AS FechaEvento,N'OC' AS TipoDocumento,@Id AS DocumentoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@RegistroRevision,N'CAMBIAR_ESTADO',
                @EstadoRevision,N'APROBADO',NULL,@MotivoAceptacion,@DetalleAceptacion;
        END;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Seguimiento_ObtenerArchivo.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguimiento_ObtenerArchivo
    @UsuarioId NVARCHAR(128),@Tipo NVARCHAR(20),@ArchivoId BIGINT, @Descargar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Permiso NVARCHAR(100)=CASE @Tipo WHEN N'OC' THEN N'OC_CLIENTE.CONSULTAR' WHEN N'GUIA' THEN N'GUIAS.CONSULTAR' WHEN N'CONFIRMACION' THEN N'CONFIRMACIONES.CONSULTAR' END;
        IF @Permiso IS NULL OR NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=@Permiso)
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para consultar este archivo.',1;
        IF @Descargar=1 AND dbo.fn_HYTPermiso(@UsuarioId,REPLACE(@Permiso,N'.CONSULTAR',N'.DESCARGAR'))=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para descargar este archivo.',1;
        SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.ClaveAlmacenamiento,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256
        FROM dbo.ArchivoAdjunto AS archivo WHERE archivo.ArchivoId=@ArchivoId AND archivo.Retirado=0
          AND ((@Tipo=N'OC' AND EXISTS(SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'GUIA' AND EXISTS(SELECT 1 FROM dbo.GuiaArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'CONFIRMACION' AND EXISTS(SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId)));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Reporte_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Reporte_Listar
    @UsuarioId NVARCHAR(128),@Reporte NVARCHAR(40),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,
    @Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@Situacion NVARCHAR(30)=NULL,
    @Estado NVARCHAR(40)=NULL,@TipoAtencion NVARCHAR(20)=NULL,@Busqueda NVARCHAR(200)=NULL,
    @PendienteGuia BIT=0,@PendienteOC BIT=0,@PendienteConfirmacion BIT=0,@PendienteFacturar BIT=0,@Exportar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Exportar=1 AND dbo.fn_HYTPermiso(@UsuarioId,N'REPORTES.IMPRIMIR')=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para imprimir reportes.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0,@VerGuia BIT=0,@VerConfirmacion BIT=0,@VerFactura BIT=0,@VerEjecucion BIT=0;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar reportes.',1;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END),
            @VerGuia=MAX(CASE WHEN permiso.CodigoPermiso=N'GUIAS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerConfirmacion=MAX(CASE WHEN permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR' THEN 1 ELSE 0 END),
            @VerFactura=MAX(CASE WHEN permiso.CodigoPermiso=N'FACTURADO.CONSULTAR' THEN 1 ELSE 0 END),
            @VerEjecucion=MAX(CASE WHEN permiso.CodigoPermiso=N'EJECUCION.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        DECLARE @Legado NVARCHAR(40)=@Reporte;
        SET @Reporte=CASE WHEN @Reporte=N'GANADOS' THEN N'PRESUPUESTO' WHEN @Reporte IN(N'OC_COMPLETADAS',N'OC_CON_ACTA',N'OC_PENDIENTES_FACTURA') THEN N'OC_CLIENTE' ELSE @Reporte END;
        IF @Reporte IS NULL OR @Reporte NOT IN(N'PRESUPUESTO',N'OC_CLIENTE') OR @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @Desde>=@HastaExclusiva
            THROW 51601,N'[HYT_VALIDACION] Revisa reporte, fechas y página.',1;
        IF (@Reporte=N'PRESUPUESTO' AND @VerPresupuesto=0) OR (@Reporte=N'OC_CLIENTE' AND @VerOC=0)
            OR (@PendienteGuia=1 AND @VerGuia=0) OR (@PendienteOC=1 AND @VerOC=0) OR (@PendienteConfirmacion=1 AND @VerConfirmacion=0) OR (@PendienteFacturar=1 AND @VerFactura=0)
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para los documentos del reporte.',1;
        IF @TipoAtencion IS NOT NULL AND @TipoAtencion NOT IN(N'INVITACION',N'EMERGENCIA') THROW 51601,N'[HYT_VALIDACION] Tipo de atención no válido.',1;
        IF @Estado IS NOT NULL AND ((@Reporte=N'PRESUPUESTO' AND @Estado NOT IN(N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'CANCELADO')) OR (@Reporte=N'OC_CLIENTE' AND @Estado NOT IN(N'REGISTRADA',N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE',N'CANCELADA'))) THROW 51601,N'[HYT_VALIDACION] Estado no válido para el reporte.',1;
        IF @Reporte=N'PRESUPUESTO' AND @Situacion IS NOT NULL THROW 51601,N'[HYT_VALIDACION] El seguimiento corresponde a OC.',1;
        IF @Situacion IS NOT NULL AND @Situacion NOT IN(N'SIN_CONFIRMAR',N'CONFIRMADAS',N'LISTAS',N'FACTURADAS') THROW 51601,N'[HYT_VALIDACION] Situación no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.PresupuestoId ELSE oc.OrdenCompraId END AS Id,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.NumeroPresupuesto ELSE CONVERT(NVARCHAR(100),oc.NumeroRegistroInterno) END AS Numero,
            p.NumeroPresupuesto AS PresupuestoNumero,p.FechaRegistro AS PresupuestoRegistro,p.TipoAtencion,
            actual.ClienteNombre,actual.TiendaNombre,actual.Estado AS EstadoPresupuesto,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Estado ELSE oc.Estado END AS Estado,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.FechaRegistro ELSE oc.FechaRegistro END AS FechaRegistro,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.CodigoMoneda ELSE fuente.CodigoMoneda END AS CodigoMoneda,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Total ELSE COALESCE(oc.Importe,fuente.Total) END AS Total,
            CASE WHEN @VerOC=1 THEN oc.NumeroRegistroInterno END AS OCInterno,CASE WHEN @VerOC=1 THEN oc.NumeroOCCliente END AS NumeroCliente,CASE WHEN @VerOC=1 THEN oc.FechaRegistro END AS OCRegistro,
            CASE WHEN @VerOC=1 THEN oc.Estado END AS EstadoOC,
            CASE WHEN @VerGuia=1 THEN guia.NumeroActa END AS Guia,CASE WHEN @VerGuia=1 AND guia.GuiaId IS NOT NULL THEN N'Guía' END AS TipoGuia,
            CASE WHEN @VerGuia=1 THEN guia.FechaRegistro END AS GuiaRegistro,CASE WHEN @VerGuia=1 THEN guia.FechaActa END AS GuiaFecha,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.EjecucionId END AS EjecucionId,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaRegistro END AS EjecucionRegistro,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaInicioReal END AS EjecucionInicio,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaTerminoReal END AS EjecucionTermino,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.NumeroConfirmacion END AS Confirmacion,CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaRegistro END AS ConfirmacionRegistro,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaConfirmacion END AS ConfirmacionFecha,
            CASE WHEN @VerFactura=1 THEN referencia.CodigoTipoDocumento END AS TipoComprobante,CASE WHEN @VerFactura=1 THEN referencia.Serie END AS Serie,
            CASE WHEN @VerFactura=1 THEN referencia.NumeroDocumento END AS NumeroComprobante,CASE WHEN @VerFactura=1 THEN referencia.FechaRegistro END AS FacturadoRegistro,
            CASE WHEN @VerFactura=1 THEN referencia.FechaDocumento END AS ComprobanteFecha,
            CONVERT(BIT,CASE WHEN @VerFactura=1 AND referencia.ReferenciaId IS NOT NULL THEN 1 ELSE 0 END) AS Facturado,
            CASE WHEN referencia.ReferenciaId IS NOT NULL THEN N'FACTURADA' WHEN confirmacion.ConfirmacionId IS NULL THEN N'SIN CONFIRMAR'
                 WHEN oc.Estado NOT IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') AND (ejecucion.Estado IS NULL OR ejecucion.Estado<>N'TERMINADA') THEN N'TRABAJO PENDIENTE'
                 WHEN guia.GuiaId IS NULL THEN N'GUÍA PENDIENTE' ELSE N'LISTA PARA MARCAR' END AS Situacion
        INTO #Reporte
        FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.RevisionId,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS actual
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1
        WHERE ((@Reporte=N'PRESUPUESTO' AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva) OR (@Reporte=N'OC_CLIENTE' AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva))
            AND (@Estado IS NULL OR (@Reporte=N'PRESUPUESTO' AND actual.Estado=@Estado) OR (@Reporte=N'OC_CLIENTE' AND oc.Estado=@Estado))
            AND (@TipoAtencion IS NULL OR p.TipoAtencion=@TipoAtencion)
            AND (@Cliente IS NULL OR actual.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR actual.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR actual.ClienteNombre LIKE @Patron ESCAPE N'~' OR actual.TiendaNombre LIKE @Patron ESCAPE N'~' OR (@VerOC=1 AND oc.NumeroOCCliente LIKE @Patron ESCAPE N'~') OR (@VerGuia=1 AND guia.NumeroActa LIKE @Patron ESCAPE N'~') OR (@VerConfirmacion=1 AND confirmacion.NumeroConfirmacion LIKE @Patron ESCAPE N'~'))
            AND (@PendienteGuia=0 OR guia.GuiaId IS NULL) AND (@PendienteOC=0 OR oc.OrdenCompraId IS NULL)
            AND (@PendienteConfirmacion=0 OR (confirmacion.ConfirmacionId IS NULL AND oc.Estado<>N'CANCELADA')) AND (@PendienteFacturar=0 OR (referencia.ReferenciaId IS NULL AND oc.Estado<>N'CANCELADA'))
            AND (@Legado<>N'GANADOS' OR actual.Estado=N'APROBADO')
            AND (@Legado<>N'OC_COMPLETADAS' OR (oc.Estado<>N'CANCELADA' AND (oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') OR ejecucion.Estado=N'TERMINADA')))
            AND (@Legado<>N'OC_CON_ACTA' OR (oc.Estado<>N'CANCELADA' AND guia.GuiaId IS NOT NULL))
            AND (@Legado<>N'OC_PENDIENTES_FACTURA' OR (oc.Estado<>N'CANCELADA' AND referencia.ReferenciaId IS NULL));
        DELETE reporte FROM #Reporte AS reporte WHERE @Situacion IS NOT NULL AND NOT
            ((@Situacion=N'SIN_CONFIRMAR' AND reporte.Situacion=N'SIN CONFIRMAR') OR (@Situacion=N'CONFIRMADAS' AND reporte.Situacion IN(N'TRABAJO PENDIENTE',N'GUÍA PENDIENTE',N'LISTA PARA MARCAR')) OR (@Situacion=N'LISTAS' AND reporte.Situacion=N'LISTA PARA MARCAR') OR (@Situacion=N'FACTURADAS' AND reporte.Situacion=N'FACTURADA'));
        SELECT COALESCE((SELECT reporte.Id,reporte.Numero,reporte.PresupuestoNumero,reporte.PresupuestoRegistro,reporte.TipoAtencion,reporte.ClienteNombre,reporte.TiendaNombre,reporte.EstadoPresupuesto,reporte.Estado,reporte.FechaRegistro,reporte.CodigoMoneda,reporte.Total,
            reporte.OCInterno,reporte.NumeroCliente,reporte.OCRegistro,reporte.EstadoOC,reporte.Guia,reporte.TipoGuia,reporte.GuiaRegistro,reporte.GuiaFecha,
            reporte.EjecucionId,reporte.EjecucionRegistro,reporte.EjecucionInicio,reporte.EjecucionTermino,reporte.Confirmacion,reporte.ConfirmacionRegistro,reporte.ConfirmacionFecha,
            reporte.TipoComprobante,reporte.Serie,reporte.NumeroComprobante,reporte.FacturadoRegistro,reporte.ComprobanteFecha,reporte.Facturado,
            CASE WHEN @VerFactura=1 AND @VerConfirmacion=1 THEN reporte.Situacion END AS Situacion
        FROM #Reporte AS reporte ORDER BY reporte.FechaRegistro DESC,reporte.Id DESC OFFSET (CASE WHEN @Exportar=1 THEN 0 ELSE (CONVERT(BIGINT,@Pagina)-1)*20 END) ROWS FETCH NEXT (CASE WHEN @Exportar=1 THEN 2147483647 ELSE 20 END) ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Reporte;
        SELECT reporte.CodigoMoneda,SUM(reporte.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Reporte AS reporte GROUP BY reporte.CodigoMoneda ORDER BY reporte.CodigoMoneda;
        DROP TABLE #Reporte;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Usuarios_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Listar
    @UsuarioId NVARCHAR(128),@Busqueda NVARCHAR(200)=NULL,@Activo BIT=NULL,@Situacion NVARCHAR(20)=NULL,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar usuarios.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR (@Situacion IS NOT NULL AND @Situacion NOT IN(N'TEMPORAL',N'BLOQUEADO',N'DISPONIBLE')) THROW 51701,N'[HYT_VALIDACION] Filtros no válidos.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT usuario.Id,usuario.Nombres,usuario.Apellidos,usuario.Email,usuario.Activo,usuario.DebeCambiarClave,usuario.UltimoIngreso,
            usuario.FechaRegistro,usuario.FechaActualizacion,usuario.ConcurrencyStamp,
            CASE WHEN usuario.LockoutEnd>SYSUTCDATETIME() THEN N'BLOQUEADO' WHEN usuario.DebeCambiarClave=1 THEN N'TEMPORAL' ELSE N'DISPONIBLE' END AS Situacion,
            COALESCE(perfiles.Nombres,N'Sin perfil') AS Perfil
        INTO #Usuarios FROM dbo.AspNetUsers AS usuario
        OUTER APPLY(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),rol.Name),N', ') AS Nombres FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.AspNetRoles AS rol ON rol.Id=asignacion.RoleId WHERE asignacion.UserId=usuario.Id) AS perfiles
        WHERE (@Activo IS NULL OR usuario.Activo=@Activo) AND (@Busqueda IS NULL OR CONCAT_WS(N' ',usuario.Nombres,usuario.Apellidos) LIKE @Patron ESCAPE N'~' OR usuario.Email LIKE @Patron ESCAPE N'~' OR perfiles.Nombres LIKE @Patron ESCAPE N'~');
        DELETE pagina FROM #Usuarios AS pagina WHERE @Situacion IS NOT NULL AND pagina.Situacion<>@Situacion;
        SELECT COALESCE((SELECT pagina.Id,pagina.Nombres,pagina.Apellidos,pagina.Email,pagina.Activo,pagina.DebeCambiarClave,pagina.UltimoIngreso,pagina.FechaRegistro,pagina.FechaActualizacion,pagina.ConcurrencyStamp,pagina.Situacion,pagina.Perfil FROM #Usuarios AS pagina ORDER BY pagina.Nombres,pagina.Apellidos,pagina.Id OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Usuarios;
        DROP TABLE #Usuarios;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Usuarios_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Obtener
    @UsuarioId NVARCHAR(128),@Id NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar usuarios.',1;
        SELECT (SELECT usuario.Id,usuario.Nombres,usuario.Apellidos,usuario.Email,usuario.Activo,usuario.DebeCambiarClave,usuario.UltimoIngreso,usuario.FechaCambioClave,usuario.LockoutEnd,usuario.AccessFailedCount,
            usuario.ConcurrencyStamp,usuario.FechaRegistro,usuario.FechaActualizacion,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,
            (SELECT TOP(1) asignacion.RoleId FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.UserId=usuario.Id ORDER BY asignacion.RoleId) AS PerfilId,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),rol.Name),N', ') FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.AspNetRoles AS rol ON rol.Id=asignacion.RoleId WHERE asignacion.UserId=usuario.Id) AS Perfil
        FROM dbo.AspNetUsers AS usuario LEFT JOIN dbo.AspNetUsers AS creador ON creador.Id=usuario.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=usuario.UsuarioActualizacion
        WHERE usuario.Id=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Perfiles_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Perfiles_Listar
    @UsuarioId NVARCHAR(128),@Busqueda NVARCHAR(200)=NULL,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51701,N'[HYT_VALIDACION] Página no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT rol.Id,rol.Name,rol.ConcurrencyStamp,rol.FechaRegistro,rol.FechaActualizacion,
            (SELECT COUNT_BIG(1) FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.RoleId=rol.Id) AS Usuarios,
            (SELECT COUNT_BIG(1) FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id) AS Permisos
        INTO #Perfiles FROM dbo.AspNetRoles AS rol WHERE @Busqueda IS NULL OR rol.Name LIKE @Patron ESCAPE N'~';
        SELECT COALESCE((SELECT pagina.Id,pagina.Name,pagina.ConcurrencyStamp,pagina.FechaRegistro,pagina.FechaActualizacion,pagina.Usuarios,pagina.Permisos FROM #Perfiles AS pagina ORDER BY pagina.Name,pagina.Id OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]');
        SELECT COUNT_BIG(1) FROM #Perfiles;DROP TABLE #Perfiles;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Perfiles_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Perfiles_Obtener
    @UsuarioId NVARCHAR(128),@Id NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        SELECT (SELECT rol.Id,rol.Name,rol.ConcurrencyStamp,rol.FechaRegistro,rol.FechaActualizacion,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,
            JSON_QUERY((SELECT permiso.CodigoPermiso AS Codigo FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id ORDER BY permiso.CodigoPermiso FOR JSON PATH)) AS Permisos,
            (SELECT COUNT_BIG(1) FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.RoleId=rol.Id) AS Usuarios
        FROM dbo.AspNetRoles AS rol LEFT JOIN dbo.AspNetUsers AS creador ON creador.Id=rol.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=rol.UsuarioActualizacion WHERE rol.Id=@Id
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Usuarios_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Guardar
    @UsuarioId NVARCHAR(128),@Datos NVARCHAR(MAX),@CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF ISJSON(@Datos)<>1 OR @CorrelacionId IS NULL THROW 51701,N'[HYT_VALIDACION] Datos no válidos.',1;
        DECLARE @Id NVARCHAR(128)=JSON_VALUE(@Datos,N'$.Id'),@Version NVARCHAR(MAX)=JSON_VALUE(@Datos,N'$.Version'),
            @Nombres NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Names'))),N''),@Apellidos NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Surnames'))),N''),
            @Email NVARCHAR(256)=LOWER(NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Email'))),N'')),@RolId NVARCHAR(128)=JSON_VALUE(@Datos,N'$.RoleId'),
            @Activo BIT=CASE JSON_VALUE(@Datos,N'$.Active') WHEN N'true' THEN 1 WHEN N'false' THEN 0 END,@Hash NVARCHAR(MAX)=JSON_VALUE(@Datos,N'$.PasswordHash'),@Nuevo BIT;
        SET @Nuevo=CASE WHEN @Version IS NULL THEN 1 ELSE 0 END;
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Nuevo=1 THEN N'USUARIOS.CREAR' ELSE N'USUARIOS.MODIFICAR' END)=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para guardar este usuario.',1;
        IF NULLIF(@Id,N'') IS NULL OR LEN(JSON_VALUE(@Datos,N'$.Id'))>128 OR @Nombres IS NULL OR @Apellidos IS NULL OR @Email IS NULL OR @Email NOT LIKE N'%_@_%._%' OR @Activo IS NULL OR LEN(JSON_VALUE(@Datos,N'$.Names'))>150 OR LEN(JSON_VALUE(@Datos,N'$.Surnames'))>150 OR LEN(JSON_VALUE(@Datos,N'$.Email'))>256 OR LEN(JSON_VALUE(@Datos,N'$.RoleId'))>128
            THROW 51701,N'[HYT_VALIDACION] Completa nombres, apellidos, correo y perfil.',1;
        IF @Nuevo=1 AND (NULLIF(@Hash,N'') IS NULL OR LEN(@Hash)>1024) THROW 51701,N'[HYT_VALIDACION] Registra una contraseña temporal válida.',1;
        IF @Nuevo=0 AND @Hash IS NOT NULL THROW 51701,N'[HYT_VALIDACION] Utiliza Restablecer contraseña para cambiar la clave de otro usuario.',1;
        IF @@TRANCOUNT<>0 THROW 51701,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:Administradores',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
        IF @Bloqueo<0 THROW 51703,N'[HYT_CONCURRENCIA] Otro administrador está actualizando accesos. Vuelve a intentarlo.',1;
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Version IS NULL THEN N'USUARIOS.CREAR' ELSE N'USUARIOS.MODIFICAR' END)=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id<>@Id AND
            (usuario.NormalizedEmail=UPPER(@Email) OR usuario.NormalizedUserName=UPPER(@Email)))
            THROW 51701,N'[HYT_VALIDACION] Ya existe un usuario con ese correo.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetRoles AS rol WITH(HOLDLOCK) WHERE rol.Id=@RolId) THROW 51701,N'[HYT_VALIDACION] El perfil seleccionado no existe.',1;
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.EDITAR')=0 AND
            (EXISTS(SELECT 1 FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=@RolId AND dbo.fn_HYTPermiso(@UsuarioId,permiso.CodigoPermiso)=0)
             OR EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId WHERE asignacion.UserId=@Id AND permiso.CodigoPermiso IN(N'USUARIOS.EDITAR',N'PERFILES.EDITAR')))
            THROW 51700,N'[HYT_PERMISO] Solo un administrador de perfiles puede gestionar cuentas con esos privilegios.',1;
        DECLARE @Anterior NVARCHAR(MAX),@Despues NVARCHAR(MAX),@Accion NVARCHAR(60),@Perfil NVARCHAR(256)=(SELECT rol.Name FROM dbo.AspNetRoles AS rol WHERE rol.Id=@RolId);
        IF @Nuevo=1
        BEGIN
            IF EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario WITH(UPDLOCK,HOLDLOCK) WHERE usuario.Id=@Id) THROW 51703,N'[HYT_CONCURRENCIA] El usuario ya fue registrado. Vuelve al listado.',1;
            INSERT INTO dbo.AspNetUsers(Id,UserName,NormalizedUserName,Email,NormalizedEmail,Nombres,Apellidos,Activo,PasswordHash,SecurityStamp,ConcurrencyStamp,DebeCambiarClave,UsuarioRegistro)
            VALUES(@Id,@Email,UPPER(@Email),@Email,UPPER(@Email),@Nombres,@Apellidos,@Activo,@Hash,CONVERT(NVARCHAR(36),NEWID()),CONVERT(NVARCHAR(36),NEWID()),1,@UsuarioId);
            SET @Accion=N'CREAR_USUARIO';
        END
        ELSE
        BEGIN
            SELECT @Anterior=(SELECT usuario.Nombres,usuario.Apellidos,usuario.Email AS Correo,usuario.Activo,usuario.DebeCambiarClave,
                (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),rol.Name),N', ') FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.AspNetRoles AS rol ON rol.Id=asignacion.RoleId WHERE asignacion.UserId=usuario.Id) AS Perfil FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.AspNetUsers AS usuario WITH(UPDLOCK,HOLDLOCK) WHERE usuario.Id=@Id AND usuario.ConcurrencyStamp=@Version;
            IF @Anterior IS NULL THROW 51703,N'[HYT_CONCURRENCIA] El usuario cambió. Vuelve a consultarlo.',1;
            IF EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id=@Id AND usuario.Activo<>@Activo) THROW 51701,N'[HYT_VALIDACION] Usa Activar o Desactivar desde la ficha del usuario.',1;
            UPDATE dbo.AspNetUsers SET UserName=@Email,NormalizedUserName=UPPER(@Email),Email=@Email,NormalizedEmail=UPPER(@Email),Nombres=@Nombres,Apellidos=@Apellidos,Activo=@Activo,
                SecurityStamp=CONVERT(NVARCHAR(36),NEWID()),ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
            DELETE asignacion FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.UserId=@Id;
            SET @Accion=N'ACTUALIZAR_USUARIO';
        END;
        INSERT INTO dbo.AspNetUserRoles(UserId,RoleId) VALUES(@Id,@RolId);
        EXEC dbo.spu_Seguridad_ProtegerAdministrador;
        SELECT @Despues=(SELECT usuario.Nombres,usuario.Apellidos,usuario.Email AS Correo,usuario.Activo,usuario.DebeCambiarClave,@Perfil AS Perfil FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.AspNetUsers AS usuario WHERE usuario.Id=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Despues) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'AspNetUsers',@Id,@Accion,NULL,NULL,NULL,NULL,@Cambios;
        COMMIT;SET @Propia=0;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Usuarios_Accion.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Accion
    @UsuarioId NVARCHAR(128),@Id NVARCHAR(128),@VersionOriginal NVARCHAR(MAX),@Accion NVARCHAR(30),@Motivo NVARCHAR(1000),@CorrelacionId UNIQUEIDENTIFIER,@PasswordHash NVARCHAR(MAX)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        DECLARE @Permiso NVARCHAR(80)=CASE @Accion WHEN N'ACTIVAR' THEN N'USUARIOS.ACTIVAR' WHEN N'DESACTIVAR' THEN N'USUARIOS.ACTIVAR' WHEN N'RESTABLECER' THEN N'USUARIOS.RESTABLECER' WHEN N'DESBLOQUEAR' THEN N'USUARIOS.DESBLOQUEAR' WHEN N'CERRAR_SESIONES' THEN N'USUARIOS.SESIONES' END;
        IF @Permiso IS NULL OR dbo.fn_HYTPermiso(@UsuarioId,@Permiso)=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción de usuario.',1;
        IF NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL OR @CorrelacionId IS NULL OR (@Accion=N'RESTABLECER' AND (NULLIF(@PasswordHash,N'') IS NULL OR LEN(@PasswordHash)>1024)) THROW 51701,N'[HYT_VALIDACION] Registra motivo y contraseña temporal cuando corresponda.',1;
        IF @@TRANCOUNT<>0 THROW 51701,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:Administradores',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
        IF @Bloqueo<0 THROW 51703,N'[HYT_CONCURRENCIA] Otro administrador está actualizando accesos. Vuelve a intentarlo.',1;
        IF dbo.fn_HYTPermiso(@UsuarioId,@Permiso)=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;

        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.EDITAR')=0 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId WHERE asignacion.UserId=@Id AND dbo.fn_HYTPermiso(@UsuarioId,permiso.CodigoPermiso)=0) THROW 51700,N'[HYT_PERMISO] Solo un administrador de perfiles puede gestionar esta cuenta.',1;
        DECLARE @Anterior NVARCHAR(MAX),@Despues NVARCHAR(MAX);
        SELECT @Anterior=(SELECT usuario.Activo,usuario.DebeCambiarClave,CASE WHEN usuario.LockoutEnd>SYSUTCDATETIME() THEN N'Bloqueado' ELSE N'Disponible' END AS Acceso FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM dbo.AspNetUsers AS usuario WITH(UPDLOCK,HOLDLOCK) WHERE usuario.Id=@Id AND usuario.ConcurrencyStamp=@VersionOriginal;
        IF @Anterior IS NULL THROW 51703,N'[HYT_CONCURRENCIA] El usuario cambió. Vuelve a consultarlo.',1;
        UPDATE dbo.AspNetUsers SET Activo=CASE @Accion WHEN N'ACTIVAR' THEN 1 WHEN N'DESACTIVAR' THEN 0 ELSE Activo END,
            PasswordHash=CASE WHEN @Accion=N'RESTABLECER' THEN @PasswordHash ELSE PasswordHash END,
            DebeCambiarClave=CASE WHEN @Accion=N'RESTABLECER' THEN 1 ELSE DebeCambiarClave END,
            LockoutEnd=CASE WHEN @Accion IN(N'RESTABLECER',N'DESBLOQUEAR') THEN NULL ELSE LockoutEnd END,
            AccessFailedCount=CASE WHEN @Accion IN(N'RESTABLECER',N'DESBLOQUEAR') THEN 0 ELSE AccessFailedCount END,
            SecurityStamp=CONVERT(NVARCHAR(36),NEWID()),ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        EXEC dbo.spu_Seguridad_ProtegerAdministrador;
        SELECT @Despues=(SELECT usuario.Activo,usuario.DebeCambiarClave,CASE WHEN usuario.LockoutEnd>SYSUTCDATETIME() THEN N'Bloqueado' ELSE N'Disponible' END AS Acceso FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) FROM dbo.AspNetUsers AS usuario WHERE usuario.Id=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Despues) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),@Evento NVARCHAR(60)=N'USUARIO_'+@Accion;
        -- El actor fue autorizado con el bloqueo de administración, antes de desactivar la cuenta.
        IF @Accion=N'DESACTIVAR' AND @UsuarioId=@Id
            INSERT INTO dbo.EventoAuditoria(CorrelacionId,Entidad,RegistroId,Accion,UsuarioId,Motivo,CambiosJson)
            VALUES(@CorrelacionId,N'AspNetUsers',@Id,@Evento,@UsuarioId,@Motivo,@Cambios);
        ELSE
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'AspNetUsers',@Id,@Evento,NULL,NULL,NULL,@Motivo,@Cambios;
        COMMIT;SET @Propia=0;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Perfiles_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Perfiles_Guardar
    @UsuarioId NVARCHAR(128),@Datos NVARCHAR(MAX),@CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.EDITAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para administrar perfiles.',1;
        IF ISJSON(@Datos)<>1 OR @CorrelacionId IS NULL THROW 51701,N'[HYT_VALIDACION] Datos no válidos.',1;
        DECLARE @Id NVARCHAR(128)=JSON_VALUE(@Datos,N'$.Id'),@Version NVARCHAR(MAX)=JSON_VALUE(@Datos,N'$.Version'),@Nombre NVARCHAR(256)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Name'))),N''),@Lista NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Permissions');
        IF NULLIF(@Id,N'') IS NULL OR LEN(JSON_VALUE(@Datos,N'$.Id'))>128 OR @Nombre IS NULL OR LEN(JSON_VALUE(@Datos,N'$.Name'))>256 OR @Lista IS NULL OR LEFT(LTRIM(@Lista),1)<>N'[' THROW 51701,N'[HYT_VALIDACION] Completa el nombre y los permisos del perfil.',1;
        IF EXISTS(SELECT 1 FROM OPENJSON(@Lista) AS entrada WHERE entrada.type<>1 OR LEN(entrada.value)>80 OR NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=entrada.value)) THROW 51701,N'[HYT_VALIDACION] El perfil incluye un permiso no válido.',1;
        DECLARE @Permisos TABLE(Codigo NVARCHAR(80) NOT NULL PRIMARY KEY);
        INSERT INTO @Permisos(Codigo) SELECT DISTINCT CONVERT(NVARCHAR(80),entrada.value) FROM OPENJSON(@Lista) AS entrada;
        IF EXISTS(SELECT 1 FROM @Permisos AS elegido INNER JOIN dbo.Permiso AS consulta ON consulta.Codigo=LEFT(elegido.Codigo,CHARINDEX(N'.',elegido.Codigo)-1)+N'.CONSULTAR' WHERE elegido.Codigo NOT LIKE N'%.CONSULTAR' AND NOT EXISTS(SELECT 1 FROM @Permisos AS actual WHERE actual.Codigo=consulta.Codigo)) THROW 51701,N'[HYT_VALIDACION] Marca Consultar antes de asignar acciones de un módulo.',1;
        INSERT INTO @Permisos(Codigo) SELECT DISTINCT LEFT(elegido.Codigo,CHARINDEX(N'.',elegido.Codigo)-1)+N'.EDITAR' FROM @Permisos AS elegido
        INNER JOIN dbo.Permiso AS legado ON legado.Codigo=LEFT(elegido.Codigo,CHARINDEX(N'.',elegido.Codigo)-1)+N'.EDITAR'
        WHERE (elegido.Codigo LIKE N'%.CREAR' OR elegido.Codigo LIKE N'%.MODIFICAR' OR elegido.Codigo LIKE N'%.ACTIVAR' OR elegido.Codigo LIKE N'%.CAMBIAR_ESTADO')
        AND NOT EXISTS(SELECT 1 FROM @Permisos AS actual WHERE actual.Codigo=legado.Codigo);
        IF @@TRANCOUNT<>0 THROW 51701,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:Administradores',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
        IF @Bloqueo<0 THROW 51703,N'[HYT_CONCURRENCIA] Otro administrador está actualizando accesos. Vuelve a intentarlo.',1;
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.EDITAR')=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WHERE rol.Id<>@Id AND rol.NormalizedName=UPPER(@Nombre))
            THROW 51701,N'[HYT_VALIDACION] Ya existe un perfil con ese nombre.',1;
        DECLARE @Anterior NVARCHAR(MAX),@Despues NVARCHAR(MAX),@Evento NVARCHAR(60);
        IF @Version IS NULL
        BEGIN
            IF EXISTS(SELECT 1 FROM dbo.AspNetRoles AS rol WITH(UPDLOCK,HOLDLOCK) WHERE rol.Id=@Id) THROW 51703,N'[HYT_CONCURRENCIA] El perfil ya fue registrado.',1;
            INSERT INTO dbo.AspNetRoles(Id,Name,NormalizedName,ConcurrencyStamp,UsuarioRegistro) VALUES(@Id,@Nombre,UPPER(@Nombre),CONVERT(NVARCHAR(36),NEWID()),@UsuarioId);
            SET @Evento=N'CREAR_PERFIL';
        END
        ELSE
        BEGIN
            SELECT @Anterior=(SELECT rol.Name AS Perfil,JSON_QUERY((SELECT permiso.CodigoPermiso AS Codigo FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id ORDER BY permiso.CodigoPermiso FOR JSON PATH)) AS Permisos FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
            FROM dbo.AspNetRoles AS rol WITH(UPDLOCK,HOLDLOCK) WHERE rol.Id=@Id AND rol.ConcurrencyStamp=@Version;
            IF @Anterior IS NULL THROW 51703,N'[HYT_CONCURRENCIA] El perfil cambió. Vuelve a consultarlo.',1;
            UPDATE dbo.AspNetRoles SET Name=@Nombre,NormalizedName=UPPER(@Nombre),ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
            DELETE permiso FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=@Id;
            SET @Evento=N'ACTUALIZAR_PERFIL';
        END;
        INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso) SELECT @Id,permiso.Codigo FROM @Permisos AS permiso;
        EXEC dbo.spu_Seguridad_ProtegerAdministrador;
        UPDATE usuario SET ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        FROM dbo.AspNetUsers AS usuario WHERE EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.UserId=usuario.Id AND asignacion.RoleId=@Id);
        SELECT @Despues=(SELECT rol.Name AS Perfil,JSON_QUERY((SELECT permiso.CodigoPermiso AS Codigo FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id ORDER BY permiso.CodigoPermiso FOR JSON PATH)) AS Permisos FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) FROM dbo.AspNetRoles AS rol WHERE rol.Id=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Despues) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'AspNetRoles',@Id,@Evento,NULL,NULL,NULL,NULL,@Cambios;
        COMMIT;SET @Propia=0;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Usuario_CambiarClave.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuario_CambiarClave
    @UsuarioId NVARCHAR(128),@SelloEsperado NVARCHAR(MAX),@HashEsperado NVARCHAR(MAX),@PasswordHash NVARCHAR(MAX),@Obligatorio BIT,@CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NULLIF(@PasswordHash,N'') IS NULL OR LEN(@PasswordHash)>1024 OR @CorrelacionId IS NULL OR @Obligatorio IS NULL THROW 51701,N'[HYT_VALIDACION] Datos de contraseña no válidos.',1;
        IF @@TRANCOUNT<>0 THROW 51701,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;SET @Propia=1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario WITH(UPDLOCK,HOLDLOCK) WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND usuario.SecurityStamp=@SelloEsperado AND usuario.PasswordHash=@HashEsperado AND usuario.DebeCambiarClave=@Obligatorio AND (usuario.LockoutEnd IS NULL OR usuario.LockoutEnd<=SYSUTCDATETIME()))
            THROW 51703,N'[HYT_CONCURRENCIA] Tu acceso cambió. Inicia sesión nuevamente.',1;
        IF @PasswordHash=@HashEsperado THROW 51701,N'[HYT_VALIDACION] La contraseña nueva debe ser diferente.',1;
        UPDATE dbo.AspNetUsers SET PasswordHash=@PasswordHash,DebeCambiarClave=0,FechaCambioClave=SYSUTCDATETIME(),SecurityStamp=CONVERT(NVARCHAR(36),NEWID()),ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@UsuarioId;
        DECLARE @Evento NVARCHAR(60)=CASE WHEN @Obligatorio=1 THEN N'CAMBIAR_CLAVE_OBLIGATORIA' ELSE N'CAMBIAR_CLAVE' END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'AspNetUsers',@UsuarioId,@Evento,NULL,NULL,NULL,NULL,NULL;
        COMMIT;SET @Propia=0;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Auditoria_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Auditoria_Listar
    @UsuarioId NVARCHAR(128), @Desde DATETIME2(3), @HastaExclusiva DATETIME2(3),
    @Busqueda NVARCHAR(200)=NULL, @ActorFiltro NVARCHAR(128)=NULL,
    @Entidad NVARCHAR(128)=NULL, @Accion NVARCHAR(60)=NULL, @Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'AUDITORIA.CONSULTAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar auditoría.',1;

        IF @Pagina < 1 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva <= @Desde
            THROW 51700,N'[HYT_VALIDACION] Revisa las fechas o la página solicitada.',1;
        IF @Entidad=N'PresupuestoRevision' SET @Entidad=N'Presupuesto';
        DECLARE @Filtro NVARCHAR(405)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(TRIM(@Busqueda),N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT evento.EventoId, evento.FechaRegistro, evento.Entidad, evento.Accion, evento.RegistroId,
            referencia.Referencia, CONCAT_WS(N' ',actor.Nombres,NULLIF(actor.Apellidos,N'')) AS Usuario
        INTO #Eventos
        FROM dbo.EventoAuditoria AS evento
        INNER JOIN dbo.AspNetUsers AS actor ON actor.Id=evento.UsuarioId
        LEFT JOIN dbo.AspNetUsers AS afectado ON evento.Entidad IN (N'Usuario', N'AspNetUsers', N'Acceso') AND afectado.Id=evento.RegistroId
        LEFT JOIN dbo.AspNetRoles AS perfil ON evento.Entidad IN (N'Perfil', N'AspNetRoles') AND perfil.Id=evento.RegistroId
        LEFT JOIN dbo.PresupuestoRevision AS revision ON evento.Entidad=N'PresupuestoRevision' AND revision.RevisionId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId=CASE
            WHEN evento.Entidad=N'Presupuesto' THEN TRY_CONVERT(BIGINT,evento.RegistroId)
            WHEN evento.Entidad=N'PresupuestoRevision' THEN revision.PresupuestoId END
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON evento.Entidad IN (N'OrdenCompraCliente',N'OC_CLIENTE') AND oc.OrdenCompraId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.GuiaConformidad AS guia ON evento.Entidad IN (N'GuiaConformidad',N'GUIA') AND guia.GuiaId=TRY_CONVERT(BIGINT,evento.RegistroId)
        CROSS APPLY (VALUES (COALESCE(afectado.Email,perfil.Name,presupuesto.NumeroPresupuesto,oc.NumeroOCCliente,guia.NumeroActa,evento.RegistroId))) AS referencia(Referencia)
        WHERE evento.FechaRegistro >= @Desde AND evento.FechaRegistro < @HastaExclusiva
            AND (@ActorFiltro IS NULL OR actor.Id=@ActorFiltro OR CHARINDEX(@ActorFiltro,CONCAT_WS(N' ',actor.Nombres,actor.Apellidos,actor.Email))>0)
            AND (@Entidad IS NULL OR evento.Entidad=@Entidad
                OR (@Entidad=N'Presupuesto' AND evento.Entidad=N'PresupuestoRevision'))
            AND (@Accion IS NULL OR evento.Accion=@Accion)
            AND (@Filtro IS NULL OR referencia.Referencia LIKE @Filtro ESCAPE N'~'
                OR evento.RegistroId LIKE @Filtro ESCAPE N'~' OR evento.Motivo LIKE @Filtro ESCAPE N'~'
                OR CONCAT_WS(N' ',actor.Nombres,actor.Apellidos,actor.Email) LIKE @Filtro ESCAPE N'~');

        SELECT (SELECT EventoId, FechaRegistro, Entidad, Accion, RegistroId, Referencia, Usuario
            FROM #Eventos ORDER BY FechaRegistro DESC, EventoId DESC
            OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Eventos;
        DROP TABLE #Eventos;
        SELECT (SELECT modulo.Entidad, evento.Accion
            FROM dbo.EventoAuditoria AS evento
            CROSS APPLY (VALUES (CASE WHEN evento.Entidad=N'PresupuestoRevision'
                THEN N'Presupuesto' ELSE evento.Entidad END)) AS modulo(Entidad)
            WHERE evento.FechaRegistro >= @Desde AND evento.FechaRegistro < @HastaExclusiva
            GROUP BY modulo.Entidad, evento.Accion ORDER BY modulo.Entidad, evento.Accion FOR JSON PATH) AS Opciones;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Auditoria_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Auditoria_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'AUDITORIA.CONSULTAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar auditoría.',1;

        SELECT (SELECT evento.EventoId, evento.FechaRegistro, evento.Entidad, evento.Accion, evento.RegistroId,
            referencia.Referencia, CONCAT_WS(N' ',actor.Nombres,NULLIF(actor.Apellidos,N'')) AS Usuario, evento.FechaEvento, evento.Motivo,
            evento.EstadoAnterior, evento.EstadoNuevo, JSON_QUERY(evento.CambiosJson) AS Cambios
        FROM dbo.EventoAuditoria AS evento
        INNER JOIN dbo.AspNetUsers AS actor ON actor.Id=evento.UsuarioId
        LEFT JOIN dbo.AspNetUsers AS afectado ON evento.Entidad IN (N'Usuario', N'AspNetUsers', N'Acceso') AND afectado.Id=evento.RegistroId
        LEFT JOIN dbo.AspNetRoles AS perfil ON evento.Entidad IN (N'Perfil', N'AspNetRoles') AND perfil.Id=evento.RegistroId
        LEFT JOIN dbo.PresupuestoRevision AS revision ON evento.Entidad=N'PresupuestoRevision' AND revision.RevisionId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId=CASE
            WHEN evento.Entidad=N'Presupuesto' THEN TRY_CONVERT(BIGINT,evento.RegistroId)
            WHEN evento.Entidad=N'PresupuestoRevision' THEN revision.PresupuestoId END
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON evento.Entidad IN (N'OrdenCompraCliente',N'OC_CLIENTE') AND oc.OrdenCompraId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.GuiaConformidad AS guia ON evento.Entidad IN (N'GuiaConformidad',N'GUIA') AND guia.GuiaId=TRY_CONVERT(BIGINT,evento.RegistroId)
        CROSS APPLY (VALUES (COALESCE(afectado.Email,perfil.Name,presupuesto.NumeroPresupuesto,oc.NumeroOCCliente,guia.NumeroActa,evento.RegistroId))) AS referencia(Referencia)
        WHERE evento.EventoId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: Script/032_PermisosAdministracion.sql
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_Seguridad_Catalogos TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Accion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuario_CambiarClave TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Auditoria_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Auditoria_Obtener TO hyt_app_executor;
GO


COMMIT TRANSACTION;
GO
