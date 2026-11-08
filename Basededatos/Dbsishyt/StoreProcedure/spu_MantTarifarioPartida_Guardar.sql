-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda TarifarioPartida con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
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
