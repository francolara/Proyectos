-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda UnidadMedida con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Expone la homologación SUNAT del maestro de unidades.
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
