-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista Cliente, búsqueda SQL y página fija de veinte.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Listar
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
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.ClienteId AS Id, m.Codigo AS Codigo, CASE WHEN m.TipoEntidad=N'NATURAL' THEN CONCAT(m.Nombres, N' ', m.Apellidos) ELSE m.RazonSocial END AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            CONCAT_WS(N' / ',NULLIF(m.DireccionFiscal,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=m.CodigoUbigeo)) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.Cliente AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR CASE WHEN m.TipoEntidad=N'NATURAL' THEN CONCAT(m.Nombres, N' ', m.Apellidos) ELSE m.RazonSocial END LIKE @Patron ESCAPE N'~' OR m.CodigoSAP LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
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
