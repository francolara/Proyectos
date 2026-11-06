"""Prepara publicación/copia transaccionales. No conecta SQL Server."""
import prepare_maintenance as base
from maintenance_schema import BY_KEY

permission = base.permission(BY_KEY['versiones-tarifario'], 'EDITAR')
start = permission + '''        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
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
'''
parameters = '    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER'
publish = start + '''        IF @Estado<>N'BORRADOR' THROW 51223,N'[HYT_VALIDACION] Solo se publica una versión borrador.',1;
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
'''
base.procedure('spu_MantTarifarioVersion_Publicar', 'Publica una versión sin alterar versiones ni presupuestos históricos.', parameters, publish, True)

clone = start + '''        SELECT @Numero=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
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
'''
base.procedure('spu_MantTarifarioVersion_Clonar', 'Copia valores y partidas a nuevo borrador con auditoría y concurrencia.', parameters, clone, True)
print('Publicación y copia preparadas en archivos; sin ejecución SQL.')
