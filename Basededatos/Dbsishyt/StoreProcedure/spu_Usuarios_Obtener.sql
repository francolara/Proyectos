-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Usuarios_Obtener; administración segura y auditada.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuarios_Obtener
    @UsuarioId NVARCHAR(128),@Id NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar usuarios.',1;
        SELECT (SELECT usuario.Id,usuario.Nombres,usuario.Apellidos,usuario.Email,usuario.Activo,usuario.DebeCambiarClave,usuario.UltimoIngreso,usuario.FechaCambioClave,usuario.LockoutEnd,usuario.AccessFailedCount,
            usuario.ConcurrencyStamp,usuario.FechaRegistro,usuario.FechaActualizacion,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,
            (SELECT TOP(1) asignacion.RoleId FROM dbo.AspNetUserRoles AS asignacion WHERE asignacion.UserId=usuario.Id ORDER BY asignacion.RoleId) AS PerfilId,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),rol.Name),N', ') FROM dbo.AspNetUserRoles AS asignacion INNER JOIN dbo.AspNetRoles AS rol ON rol.Id=asignacion.RoleId WHERE asignacion.UserId=usuario.Id) AS Perfil
        FROM dbo.AspNetUsers AS usuario LEFT JOIN dbo.AspNetUsers AS creador ON creador.Id=usuario.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=usuario.UsuarioActualizacion
        WHERE usuario.Id=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
