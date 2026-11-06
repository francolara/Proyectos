-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda TarifarioVersion con permisos, concurrencia, referencias y auditoría.
-- =============================================
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
