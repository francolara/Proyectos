-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Recupera un adjunto vigente solo con permiso del documento vinculado.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguimiento_ObtenerArchivo
    @UsuarioId NVARCHAR(128),@Tipo NVARCHAR(20),@ArchivoId BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Permiso NVARCHAR(100)=CASE @Tipo WHEN N'OC' THEN N'OC_CLIENTE.CONSULTAR' WHEN N'GUIA' THEN N'GUIAS.CONSULTAR' WHEN N'CONFIRMACION' THEN N'CONFIRMACIONES.CONSULTAR' END;
        IF @Permiso IS NULL OR NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=@Permiso)
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para consultar este archivo.',1;
        SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.ClaveAlmacenamiento,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256
        FROM dbo.ArchivoAdjunto AS archivo WHERE archivo.ArchivoId=@ArchivoId AND archivo.Retirado=0
          AND ((@Tipo=N'OC' AND EXISTS(SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'GUIA' AND EXISTS(SELECT 1 FROM dbo.GuiaArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'CONFIRMACION' AND EXISTS(SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId)));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO