-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Publica una versión sin alterar versiones ni presupuestos históricos.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Publicar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
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
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT, @Numero INT, @Estado NVARCHAR(20);
        SELECT @TarifarioId=v.TarifarioId FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        IF @TarifarioId IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión ya no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS t WITH (UPDLOCK,HOLDLOCK) WHERE t.TarifarioId=@TarifarioId AND t.Activo=1)
            THROW 51222,N'[HYT_VALIDACION] El tarifario no está activo.',1;
        SELECT @Estado=v.Estado FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK)
        WHERE v.VersionId=@Id AND v.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión cambió. Vuelve a consultarla.',1;
        IF @Estado<>N'BORRADOR' THROW 51223,N'[HYT_VALIDACION] Solo se publica una versión borrador.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id AND p.Activo=1)
            THROW 51224,N'[HYT_VALIDACION] Agrega al menos una partida activa antes de publicar.',1;
        IF EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK)
            INNER JOIN dbo.ServicioMaterial AS material WITH (HOLDLOCK) ON material.MaterialId=p.MaterialId
            INNER JOIN dbo.UnidadMedida AS unidad WITH (HOLDLOCK) ON unidad.Id=material.UnidadMedidaId
            WHERE p.VersionId=@Id AND p.Activo=1 AND (p.PrecioPEN IS NULL OR material.Activo=0 OR unidad.Activo=0))
            THROW 51225,N'[HYT_VALIDACION] Revisa precios PEN, materiales y unidades de las partidas activas.',1;
        DECLARE @Anterior NVARCHAR(MAX), @Nuevo NVARCHAR(MAX), @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        SELECT @Anterior=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        UPDATE dbo.TarifarioVersion SET Estado=N'VIGENTE',FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@Id;
        SELECT @Nuevo=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'PUBLICAR',N'BORRADOR',N'VIGENTE',NULL,NULL,@Cambios;
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
