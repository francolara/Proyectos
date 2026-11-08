-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta privada de identidad, roles y permisos para acceso web.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Nombre completo y obligación de cambio de clave, sin exigir confirmación de correo.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConsultarAcceso
    @UsuarioId NVARCHAR(128) = NULL,
    @NombreNormalizado NVARCHAR(256) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF (@UsuarioId IS NULL AND @NombreNormalizado IS NULL) OR (@UsuarioId IS NOT NULL AND @NombreNormalizado IS NOT NULL)
            THROW 51100, N'Indique un unico criterio de identidad.', 1;
        DECLARE @Id NVARCHAR(128);
        SELECT @Id = usuario.Id FROM dbo.AspNetUsers AS usuario
        WHERE (@UsuarioId IS NOT NULL AND usuario.Id = @UsuarioId)
            OR (@NombreNormalizado IS NOT NULL AND usuario.NormalizedUserName = @NombreNormalizado);

        SELECT usuario.Id, usuario.UserName, LTRIM(RTRIM(usuario.Nombres+N' '+usuario.Apellidos)) AS Nombres, usuario.PasswordHash,
            COALESCE(usuario.SecurityStamp, N'') AS SecurityStamp, usuario.Activo, usuario.LockoutEnd,usuario.DebeCambiarClave
        FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @Id;
        SELECT rol.Name FROM dbo.AspNetUserRoles AS asignacion
        INNER JOIN dbo.AspNetRoles AS rol ON rol.Id = asignacion.RoleId
        WHERE asignacion.UserId = @Id AND rol.Name IS NOT NULL ORDER BY rol.Name;
        SELECT DISTINCT permiso.CodigoPermiso FROM dbo.AspNetUserRoles AS asignacion
        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
        WHERE asignacion.UserId = @Id ORDER BY permiso.CodigoPermiso;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
