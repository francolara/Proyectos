-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Permiso efectivo para operaciones administrativas, denegando cuentas temporales.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_HYTPermiso(@UsuarioId NVARCHAR(128),@Codigo NVARCHAR(80))
RETURNS BIT
AS
BEGIN
    DECLARE @Resultado BIT=0;
    IF EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId=usuario.Id
        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId
        WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND usuario.DebeCambiarClave=0 AND permiso.CodigoPermiso=@Codigo)
        SET @Resultado=1;
    RETURN @Resultado;
END;
GO
