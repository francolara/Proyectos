-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Usuarios_Guardar; permisos, concurrencia y auditoría atómica sin secretos.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Revalida autorización bajo bloqueo, detecta correos duplicados y reserva el estado para su acción auditada.
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
