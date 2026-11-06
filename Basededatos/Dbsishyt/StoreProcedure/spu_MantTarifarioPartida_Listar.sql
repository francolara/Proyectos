-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista TarifarioPartida, búsqueda SQL y página fija de veinte.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioPartida_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.PartidaId AS Id, COALESCE(m.Posicion, CONVERT(NVARCHAR(40), m.PartidaId)) AS Codigo, material.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), z.Descripcion), N', ') FROM dbo.TarifarioPartidaZona AS pz INNER JOIN dbo.Zona AS z ON z.Id = pz.ZonaId WHERE pz.PartidaId = m.PartidaId) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.TarifarioPartida AS m INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId = m.MaterialId INNER JOIN dbo.TarifarioVersion AS padre ON padre.VersionId = m.VersionId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR COALESCE(m.Posicion, CONVERT(NVARCHAR(40), m.PartidaId)) LIKE @Patron ESCAPE N'~' OR material.Descripcion LIKE @Patron ESCAPE N'~' OR m.CodigoMaterialSAP LIKE @Patron ESCAPE N'~' OR material.CodigoInterno LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND (@PadreId IS NULL OR m.VersionId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
