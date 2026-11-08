-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Copia valores y partidas a nuevo borrador con auditoría y concurrencia.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Comprueba el permiso específico del perfil antes de realizar la acción.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Clonar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.CREAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
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
        SELECT @Numero=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
        INSERT INTO dbo.TarifarioVersion (TarifarioId,NumeroVersion,VigenciaDesde,VigenciaHasta,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
        SELECT v.TarifarioId,@Numero,v.VigenciaDesde,v.VigenciaHasta,v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,N'BORRADOR',@UsuarioId
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @NuevaVersion BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY());
        INSERT INTO dbo.TarifarioVersionZona (VersionId,ZonaId)
        SELECT @NuevaVersion,z.ZonaId FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=@Id;
        DECLARE @Mapa TABLE (AnteriorId BIGINT NOT NULL PRIMARY KEY,NuevoId BIGINT NOT NULL);
        -- MERGE de solo inserción permite relacionar el identificador fuente con el generado.
        MERGE dbo.TarifarioPartida AS destino
        USING (SELECT p.PartidaId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo
               FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id) AS fuente ON 1=0
        WHEN NOT MATCHED THEN INSERT (VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,Posicion,CodigoMaterialSAP,TipoServicioId,NumeroNecesidad,Activo,UsuarioRegistro)
        VALUES (@NuevaVersion,fuente.MaterialId,fuente.CodigoMoneda,fuente.CodigoTipoIGV,fuente.PrecioPEN,fuente.PrecioUSD,fuente.Posicion,fuente.CodigoMaterialSAP,fuente.TipoServicioId,fuente.NumeroNecesidad,fuente.Activo,@UsuarioId)
        OUTPUT fuente.PartidaId,inserted.PartidaId INTO @Mapa (AnteriorId,NuevoId);
        INSERT INTO dbo.TarifarioPartidaZona (PartidaId,ZonaId)
        SELECT mapa.NuevoId,zona.ZonaId FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartidaZona AS zona ON zona.PartidaId=mapa.AnteriorId;
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoNuevo,UsuarioId,CambiosJson)
        SELECT @CorrelacionId,N'TarifarioPartida',CONVERT(NVARCHAR(100),mapa.NuevoId),N'COPIAR',CASE p.Activo WHEN 1 THEN N'true' ELSE N'false' END,@UsuarioId,
            (SELECT mapa.AnteriorId AS PartidaOrigen,p.VersionId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=p.PartidaId FOR JSON PATH)) AS Zonas
             FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartida AS p ON p.PartidaId=mapa.NuevoId;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT @Id AS VersionOrigen,@NuevaVersion AS VersionNueva,@Numero AS NumeroVersion,
            JSON_QUERY((SELECT v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,v.VigenciaDesde,v.VigenciaHasta,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=v.VersionId FOR JSON PATH)) AS Zonas
             FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@NuevaVersion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Valores,
            (SELECT COUNT_BIG(1) FROM @Mapa AS mapa) AS PartidasCopiadas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@NuevaVersion);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'COPIAR',NULL,N'BORRADOR',NULL,NULL,@Cambios;
        SET @Registro=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'GENERAR_COPIA',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @NuevaVersion AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
