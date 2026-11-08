-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Usuarios_Listar; administración segura y auditada.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Búsqueda por nombre completo y alias explícito para el resultado paginado.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Listar
    @UsuarioId NVARCHAR(128),@Busqueda NVARCHAR(200)=NULL,@Activo BIT=NULL,@Situacion NVARCHAR(20)=NULL,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar usuarios.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR (@Situacion IS NOT NULL AND @Situacion NOT IN(N'TEMPORAL',N'BLOQUEADO',N'DISPONIBLE')) THROW 51701,N'[HYT_VALIDACION] Filtros no válidos.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT usuario.Id,usuario.Nombres,usuario.Apellidos,usuario.Email,usuario.Activo,usuario.DebeCambiarClave,usuario.UltimoIngreso,
            usuario.FechaRegistro,usuario.FechaActualizacion,usuario.ConcurrencyStamp,
            CASE WHEN usuario.LockoutEnd>SYSUTCDATETIME() THEN N'BLOQUEADO' WHEN usuario.DebeCambiarClave=1 THEN N'TEMPORAL' ELSE N'DISPONIBLE' END AS Situacion,
            COALESCE(perfiles.Nombres,N'Sin perfil') AS Perfil
        INTO #Usuarios FROM dbo.AspNetUsers AS usuario
        OUTER APPLY(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),rol.Name),N', ') AS Nombres FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.AspNetRoles AS rol ON rol.Id=asignacion.RoleId WHERE asignacion.UserId=usuario.Id) AS perfiles
        WHERE (@Activo IS NULL OR usuario.Activo=@Activo) AND (@Busqueda IS NULL OR CONCAT_WS(N' ',usuario.Nombres,usuario.Apellidos) LIKE @Patron ESCAPE N'~' OR usuario.Email LIKE @Patron ESCAPE N'~' OR perfiles.Nombres LIKE @Patron ESCAPE N'~');
        DELETE pagina FROM #Usuarios AS pagina WHERE @Situacion IS NOT NULL AND pagina.Situacion<>@Situacion;
        SELECT COALESCE((SELECT pagina.Id,pagina.Nombres,pagina.Apellidos,pagina.Email,pagina.Activo,pagina.DebeCambiarClave,pagina.UltimoIngreso,pagina.FechaRegistro,pagina.FechaActualizacion,pagina.ConcurrencyStamp,pagina.Situacion,pagina.Perfil FROM #Usuarios AS pagina ORDER BY pagina.Nombres,pagina.Apellidos,pagina.Id OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Usuarios;
        DROP TABLE #Usuarios;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
