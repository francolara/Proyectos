-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Perfiles_Guardar; permisos, concurrencia y auditoría atómica sin secretos.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Aplica permisos al siguiente acceso sin invalidar la sesión y evita perfiles duplicados.
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
