-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Tarifarios unificados con búsqueda por material/SAP, edición autorizada, fechas propias y veinte filas por página.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MaterialTarifas_Listar
    @UsuarioId NVARCHAR(128),@MaterialId BIGINT,@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'MATERIALES.CONSULTAR')
            OR NOT EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'TARIFARIOS.CONSULTAR')
            THROW 51200,N'[HYT_PERMISO] Necesita permiso de consulta de materiales y tarifarios.',1;
        IF @MaterialId<0 OR @Pagina<1 OR @Pagina>100000000 THROW 51201,N'[HYT_VALIDACION] Material o página no válido.',1;
        SELECT COALESCE((SELECT partida.PartidaId AS Id,partida.MaterialId,material.CodigoInterno AS MaterialCodigo,material.Descripcion AS MaterialDescripcion,version.VersionId,tarifa.Nombre AS Tarifario,version.NumeroVersion,version.Estado,
            partida.Posicion,partida.CodigoMaterialSAP,partida.CodigoMoneda,partida.CodigoTipoIGV,partida.PrecioPEN,partida.PrecioUSD,partida.Activo,CAST(CASE WHEN tarifa.EsTarifarioMaterial=1 OR version.Estado=N'BORRADOR' THEN 1 ELSE 0 END AS BIT) AS Editable,
            tipo.Descripcion AS TipoServicio,partida.FechaRegistro,creador.Nombres AS UsuarioRegistro,partida.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=partida.PartidaId) AS Zonas
            FROM dbo.TarifarioPartida AS partida INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
            INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId LEFT JOIN dbo.TipoServicio AS tipo ON tipo.Id=partida.TipoServicioId
            INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=partida.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=partida.UsuarioActualizacion
            WHERE (@MaterialId=0 OR partida.MaterialId=@MaterialId) AND (@Busqueda IS NULL OR material.CodigoInterno LIKE N'%'+@Busqueda+N'%' OR material.Descripcion LIKE N'%'+@Busqueda+N'%' OR partida.CodigoMaterialSAP LIKE N'%'+@Busqueda+N'%') ORDER BY partida.FechaRegistro DESC,partida.PartidaId DESC
            OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM dbo.TarifarioPartida AS partida INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId WHERE (@MaterialId=0 OR partida.MaterialId=@MaterialId) AND (@Busqueda IS NULL OR material.CodigoInterno LIKE N'%'+@Busqueda+N'%' OR material.Descripcion LIKE N'%'+@Busqueda+N'%' OR partida.CodigoMaterialSAP LIKE N'%'+@Busqueda+N'%');
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifas_Listar TO hyt_app_executor;
GO
