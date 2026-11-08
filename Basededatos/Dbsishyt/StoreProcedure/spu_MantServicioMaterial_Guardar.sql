-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda ServicioMaterial con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
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
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.CodigoInterno, anterior.CodigoMaterialSAP, anterior.Descripcion, anterior.UnidadMedidaId, anterior.EspecialidadId, anterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=anterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ServicioMaterial AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.MaterialId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'MATERIALES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
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
            INSERT INTO dbo.ServicioMaterial (CodigoInterno, CodigoMaterialSAP, Descripcion, UnidadMedidaId, EspecialidadId, Activo, UsuarioRegistro)
            VALUES (@CodigoInterno, @CodigoMaterialSAP, @Descripcion, @UnidadMedidaId, @EspecialidadId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ServicioMaterial SET CodigoInterno=@CodigoInterno, CodigoMaterialSAP=@CodigoMaterialSAP, Descripcion=@Descripcion, UnidadMedidaId=@UnidadMedidaId, EspecialidadId=@EspecialidadId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE MaterialId=@Id;
        END;
        DELETE FROM dbo.MaterialZona WHERE MaterialId=@Id;
        INSERT INTO dbo.MaterialZona (MaterialId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        SELECT @Posterior=(SELECT posterior.CodigoInterno, posterior.CodigoMaterialSAP, posterior.Descripcion, posterior.UnidadMedidaId, posterior.EspecialidadId, posterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=posterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
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
