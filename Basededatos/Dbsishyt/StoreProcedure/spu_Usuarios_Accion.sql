-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Usuarios_Accion; permisos, concurrencia y auditoría atómica sin secretos.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Revalida autorización bajo bloqueo, protege cuentas con permisos superiores y registra la desactivación propia autorizada.
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
