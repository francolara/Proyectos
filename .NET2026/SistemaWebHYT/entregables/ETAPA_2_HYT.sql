-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Entrega HYT etapa 2: acceso y permisos para ejecucion manual.
-- =============================================
-- Revisar LEEME_DESPLIEGUE.md. Este archivo no se ejecuta automaticamente.
-- Aplicar despues de la etapa 1; no contiene usuarios ni contrasenas.

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

        SELECT usuario.Id, usuario.UserName, usuario.Nombres, usuario.PasswordHash,
            COALESCE(usuario.SecurityStamp, N'') AS SecurityStamp, usuario.Activo, usuario.LockoutEnd
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
                    PasswordHash = COALESCE(@HashActualizado, PasswordHash)
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


-- Fuente: StoreProcedure/spu_RegistrarSalida.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RegistrarSalida
    @UsuarioId NVARCHAR(128),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId)
        SELECT @CorrelacionId, N'AspNetUsers', usuario.Id, N'CIERRE_SESION', usuario.Id
        FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_CrearAdministradorInicial.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CrearAdministradorInicial
    @UsuarioId NVARCHAR(128),
    @Usuario NVARCHAR(256),
    @NombreNormalizado NVARCHAR(256),
    @Nombre NVARCHAR(150),
    @PasswordHash NVARCHAR(MAX),
    @SecurityStamp NVARCHAR(MAX),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51102, N'El aprovisionamiento administra su propia transaccion.', 1;
        IF NULLIF(@Usuario, N'') IS NULL OR NULLIF(@Nombre, N'') IS NULL OR NULLIF(@PasswordHash, N'') IS NULL OR NULLIF(@SecurityStamp, N'') IS NULL
            THROW 51103, N'Datos de aprovisionamiento incompletos.', 1;
        BEGIN TRANSACTION;
        SET @Propia = 1;
        DECLARE @BloqueoBootstrap INT;
        EXEC @BloqueoBootstrap = sys.sp_getapplock @Resource = N'HYT:BootstrapAdmin', @LockMode = N'Exclusive', @LockOwner = N'Transaction', @LockTimeout = 10000;
        IF @BloqueoBootstrap < 0 THROW 51109, N'No se pudo bloquear el aprovisionamiento inicial.', 1;
        IF EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK)
            INNER JOIN dbo.AspNetUserRoles AS asignacion WITH (UPDLOCK, HOLDLOCK) ON asignacion.UserId = usuario.Id
            WHERE asignacion.RoleId = N'HYT-ADMIN' AND usuario.Activo = 1
        )
            THROW 51104, N'Ya existe un administrador activo; use el mantenimiento autorizado.', 1;
        INSERT INTO dbo.AspNetUsers (Id, UserName, NormalizedUserName, Nombres, PasswordHash,
            SecurityStamp, ConcurrencyStamp, Activo, LockoutEnabled)
        VALUES (@UsuarioId, @Usuario, @NombreNormalizado, @Nombre, @PasswordHash,
            @SecurityStamp, CONVERT(NVARCHAR(36), NEWID()), 1, 1);
        INSERT INTO dbo.AspNetUserRoles (UserId, RoleId) VALUES (@UsuarioId, N'HYT-ADMIN');
        DECLARE @Detalle NVARCHAR(MAX) = (SELECT ORIGINAL_LOGIN() AS LoginDespliegue FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId, CambiosJson)
        VALUES (@CorrelacionId, N'AspNetUsers', @UsuarioId, N'ADMINISTRADOR_INICIAL', @UsuarioId, @Detalle);
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
            WHERE usuario.Id = @ActorId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'USUARIOS.EDITAR'
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
            FechaActualizacion = SYSUTCDATETIME() WHERE Id = @UsuarioId;
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


-- Fuente: Script/010_PermisosAcceso.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_ConsultarAcceso TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_FinalizarAcceso TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_RegistrarSalida TO hyt_app_executor;
-- No se concede CREAR_ADMIN ni RESTABLECER al rol web en esta etapa.
-- Las herramientas administrativas usan la cuenta de despliegue autorizada.
GO

