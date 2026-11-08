-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Perfiles_Obtener; administración segura y auditada.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Perfiles_Obtener
    @UsuarioId NVARCHAR(128),@Id NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        SELECT (SELECT rol.Id,rol.Name,rol.ConcurrencyStamp,rol.FechaRegistro,rol.FechaActualizacion,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,
            JSON_QUERY((SELECT permiso.CodigoPermiso AS Codigo FROM dbo.RolPermiso AS permiso WHERE permiso.RoleId=rol.Id ORDER BY permiso.CodigoPermiso FOR JSON PATH)) AS Permisos,
            (SELECT COUNT_BIG(1) FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.RoleId=rol.Id) AS Usuarios
        FROM dbo.AspNetRoles AS rol LEFT JOIN dbo.AspNetUsers AS creador ON creador.Id=rol.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=rol.UsuarioActualizacion WHERE rol.Id=@Id
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
