-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda ContactoCliente con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
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
