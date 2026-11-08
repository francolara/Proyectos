-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda Tienda con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
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
