-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Seguridad_Catalogos; seguridad, concurrencia y auditoría según operación.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguridad_Catalogos
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'USUARIOS.CONSULTAR')=0 AND dbo.fn_HYTPermiso(@UsuarioId,N'PERFILES.CONSULTAR')=0 THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar perfiles.',1;
        SELECT (SELECT rol.Id,rol.Name FROM dbo.AspNetRoles AS rol ORDER BY rol.Name FOR JSON PATH) AS Perfiles;
        SELECT (SELECT permiso.Codigo,permiso.Descripcion FROM dbo.Permiso AS permiso ORDER BY permiso.Codigo FOR JSON PATH) AS Permisos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
