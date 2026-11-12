-- Firma: FRANCO LARA | Create date: 10/11/2026 | Selecciona una tarifa por material, con prioridad al tarifario directo y compatibilidad con partidas históricas.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Ayuda de tarifas publicadas por cliente/fecha y moneda; zonas solo informativas.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ayuda filtrada por material y tarifa para el detalle editable del presupuesto.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Muestra el nombre del tipo de servicio asociado a la tarifa.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_AyudaTarifario
    @UsuarioId NVARCHAR(128),@ClienteId BIGINT,@Fecha DATE,@CodigoMoneda NVARCHAR(3)=N'PEN',@Busqueda NVARCHAR(200)=NULL,@Pagina INT=1,@MaterialId BIGINT=NULL,@PartidaId BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Fecha IS NULL OR @CodigoMoneda NOT IN (N'PEN',N'USD') THROW 51301,N'[HYT_VALIDACION] Revisa página, fecha y moneda.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        DECLARE @TarifaMaterialId BIGINT;
        IF @MaterialId IS NOT NULL AND @PartidaId IS NULL
            SELECT TOP (1) @TarifaMaterialId=partida.PartidaId
            FROM dbo.TarifarioPartida AS partida
            INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
            WHERE partida.MaterialId=@MaterialId
            ORDER BY tarifa.EsTarifarioMaterial DESC, partida.Activo DESC,
                CASE WHEN version.Estado=N'VIGENTE' THEN 0 ELSE 1 END,
                version.VigenciaDesde DESC, partida.PartidaId DESC;

        SELECT partida.PartidaId AS PartidaTarifarioId,material.MaterialId,material.CodigoInterno,COALESCE(partida.Posicion,material.Posicion) AS Posicion,
            COALESCE(partida.CodigoMaterialSAP,material.CodigoMaterialSAP) AS CodigoMaterialSAP,material.Descripcion,unidad.Codigo AS Unidad,
            tarifa.Codigo AS CodigoTarifario,version.NumeroVersion,partida.CodigoTipoIGV,@CodigoMoneda AS CodigoMoneda,
            CONVERT(DECIMAL(19,2),CASE WHEN @CodigoMoneda=N'USD' THEN partida.PrecioUSD ELSE partida.PrecioPEN END) AS Precio,
            servicio.Descripcion AS TipoServicio,partida.NumeroNecesidad,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=partida.PartidaId) AS Zonas,
            ROW_NUMBER() OVER(PARTITION BY partida.MaterialId ORDER BY tarifa.EsTarifarioMaterial DESC,
                CASE WHEN tarifa.ClienteId=@ClienteId THEN 0 ELSE 1 END,
                version.VigenciaDesde DESC,partida.PartidaId DESC) AS Prioridad
        INTO #Pagina FROM dbo.TarifarioPartida AS partida
        INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
        INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
        INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId
        INNER JOIN dbo.UnidadMedida AS unidad ON unidad.Id=material.UnidadMedidaId
        LEFT JOIN dbo.TipoServicio AS servicio ON servicio.Id=partida.TipoServicioId
        WHERE (@MaterialId IS NULL OR material.MaterialId=@MaterialId) AND (@PartidaId IS NULL OR partida.PartidaId=@PartidaId) AND partida.Activo=1 AND tarifa.Activo=1 AND material.Activo=1 AND unidad.Activo=1 AND version.Estado=N'VIGENTE'
            AND (@MaterialId IS NULL OR @PartidaId IS NOT NULL OR partida.PartidaId=@TarifaMaterialId)
            AND (@TarifaMaterialId IS NOT NULL OR tarifa.ClienteId IS NULL OR tarifa.ClienteId=@ClienteId)
            AND (version.VigenciaDesde IS NULL OR version.VigenciaDesde<=@Fecha) AND (version.VigenciaHasta IS NULL OR version.VigenciaHasta>=@Fecha)
            AND (@Busqueda IS NULL OR material.CodigoInterno LIKE @Patron ESCAPE N'~' OR material.Descripcion LIKE @Patron ESCAPE N'~' OR partida.CodigoMaterialSAP LIKE @Patron ESCAPE N'~' OR partida.Posicion LIKE @Patron ESCAPE N'~');
        SELECT pagina.PartidaTarifarioId,pagina.MaterialId,pagina.CodigoInterno,pagina.Posicion,pagina.CodigoMaterialSAP,pagina.Descripcion,pagina.Unidad,pagina.CodigoTarifario,pagina.NumeroVersion,
            pagina.CodigoTipoIGV,pagina.CodigoMoneda,pagina.Precio,pagina.TipoServicio,pagina.NumeroNecesidad,pagina.Zonas
        FROM #Pagina AS pagina WHERE pagina.Prioridad=1 ORDER BY pagina.CodigoInterno,pagina.PartidaTarifarioId OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina WHERE pagina.Prioridad=1;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
