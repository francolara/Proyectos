-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Perfiles_Listar; administración segura y auditada.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Alias separado para la página de resultados.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Perfiles_Listar
    @UsuarioId NVARCHAR(128),@Busqueda NVARCHAR(200)=NULL,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51701,N'[HYT_VALIDACION] Página no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT rol.Id,rol.Name,rol.ConcurrencyStamp,rol.FechaRegistro,rol.FechaActualizacion,
            (SELECT COUNT_BIG(1) FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.RoleId=rol.Id) AS Usuarios,
            (SELECT COUNT_BIG(1) FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id) AS Permisos
        INTO #Perfiles FROM dbo.AspNetRoles AS rol WHERE @Busqueda IS NULL OR rol.Name LIKE @Patron ESCAPE N'~';
        SELECT COALESCE((SELECT pagina.Id,pagina.Name,pagina.ConcurrencyStamp,pagina.FechaRegistro,pagina.FechaActualizacion,pagina.Usuarios,pagina.Permisos FROM #Perfiles AS pagina ORDER BY pagina.Name,pagina.Id OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]');
        SELECT COUNT_BIG(1) FROM #Perfiles;DROP TABLE #Perfiles;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
